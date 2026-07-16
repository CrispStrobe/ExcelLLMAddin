Attribute VB_Name = "modPane"
' modPane: logic for the modeless "task pane" UserForm (frmLLMPane).
'
' Pure VBA => works on locked-down Office where the Office.js web add-in is
' blocked by policy. Floating modeless panel positioned at the right edge of the
' Excel window (a true docked Custom Task Pane needs a COM/VSTO add-in). The
' form's event handlers are one-liners that delegate here, so all behaviour lives
' in this text module (source-controlled + editable).
'
' Controls expected on frmLLMPane:
'   cboProvider, cboModel, btnLoadModels, txtKey, txtBaseUrl, txtMcp, txtImgKey,
'   txtSystem, btnSave, btnTest, btnUsageReset, lblUsage, cboPreset, txtPrompt,
'   btnRunPrompt, btnRunAgent, chkAuto, lblStatus, txtOutput
Option Explicit

Private Const PANE_MARGIN As Long = 12

Public Function PaneProviders() As Variant
    PaneProviders = Array("ollama", "openai", "mistral", "nebius", "scaleway", _
                          "openrouter", "groq", "together", "cerebras", _
                          "gemini", "cohere", "huggingface", "requesty")
End Function

Public Function PanePresets() As Variant
    PanePresets = Array( _
        "-- presets --", _
        "Summarize the selected range in 3 bullet points", _
        "In D1 put the sum of B2:B10, then bold any cell over 100", _
        "Add a column classifying each row of my selection as high or low", _
        "Translate the selected cells to German", _
        "Clean and standardize the selected data (trim, title-case names)", _
        "Explain the formula in the active cell")
End Function

' Public entry point: show the pane at the right edge of the Excel window.
Public Sub ShowLLMPane()
    On Error Resume Next
    LoadConfig
    With frmLLMPane
        .StartUpPosition = 0
        .Left = Application.Left + Application.Width - .Width - PANE_MARGIN
        .Top = Application.Top + 80
        .Show vbModeless
    End With
End Sub

Public Sub Pane_Init(frm As Object)
    On Error Resume Next
    LoadConfig
    Dim p As Variant
    frm.cboProvider.Clear
    For Each p In PaneProviders()
        frm.cboProvider.AddItem p
    Next p
    frm.cboPreset.Clear
    For Each p In PanePresets()
        frm.cboPreset.AddItem p
    Next p
    frm.cboPreset.ListIndex = 0

    frm.cboProvider.Value = IIf(CurrentProvider = "", "ollama", CurrentProvider)
    frm.cboModel.Text = CurrentModel
    frm.txtKey.Text = GetAPIKey(CStr(frm.cboProvider.Value))
    frm.txtBaseUrl.Text = GetBaseURL(CStr(frm.cboProvider.Value))
    frm.txtMcp.Text = modMcp.McpUrl
    frm.txtImgKey.Text = GetAPIKey("bfl")
    frm.txtOutput.Text = ""
    PaneRefreshUsage frm
    PaneStatus frm, "Ready - provider: " & frm.cboProvider.Value
End Sub

Public Sub Pane_ProviderChanged(frm As Object)
    On Error Resume Next
    Dim prov As String: prov = LCase$(Trim$(CStr(frm.cboProvider.Value)))
    frm.txtKey.Text = GetAPIKey(prov)
    frm.txtBaseUrl.Text = GetBaseURL(prov)
    PaneStatus frm, "Provider: " & prov & IIf(prov = "ollama", " (local, no key)", "")
End Sub

Public Sub Pane_PresetChosen(frm As Object)
    On Error Resume Next
    If frm.cboPreset.ListIndex <= 0 Then Exit Sub
    frm.txtPrompt.Text = CStr(frm.cboPreset.Value)
End Sub

Public Sub Pane_LoadModels(frm As Object)
    On Error GoTo Fail
    ApplyPaneConfig frm
    Dim prov As String: prov = LCase$(Trim$(CStr(frm.cboProvider.Value)))
    PaneStatus frm, "Loading models for " & prov & "..."
    ' Live first (keyless for OpenRouter/Ollama, with-key for the rest); fall back
    ' to the curated catalog so the dropdown is never empty.
    Dim s As String, src As String
    s = LIST_MODELS(prov)
    If Left$(s, 6) = "Error:" Or Trim$(s) = "" Then
        s = DefaultModels(prov): src = "defaults"
    Else
        src = "live"
    End If
    frm.cboModel.Clear
    Dim arr() As String, i As Long, n As Long
    arr = Split(s, vbCrLf)
    For i = LBound(arr) To UBound(arr)
        If Trim$(arr(i)) <> "" Then frm.cboModel.AddItem Trim$(arr(i)): n = n + 1
    Next i
    If n = 0 Then PaneStatus frm, "No models for " & prov Else PaneStatus frm, n & " models (" & src & ")"
    Exit Sub
Fail:
    PaneStatus frm, "Model load error: " & Err.Description
End Sub

Public Sub Pane_Save(frm As Object)
    On Error GoTo Fail
    ApplyPaneConfig frm
    SaveConfig
    PaneStatus frm, "Saved: " & CurrentProvider & " / " & CurrentModel
    Exit Sub
Fail:
    PaneStatus frm, "Save error: " & Err.Description
End Sub

Public Sub Pane_Test(frm As Object)
    On Error GoTo Fail
    ApplyPaneConfig frm
    PaneStatus frm, "Testing " & CurrentProvider & "..."
    Dim r As String
    r = ChatComplete("You are a connection test.", "Reply with the single word OK.", _
                     CurrentProvider, CurrentModel)
    PaneOut frm, "[TEST " & CurrentProvider & "] " & r
    PaneRefreshUsage frm
    PaneStatus frm, IIf(Left$(r, 6) = "Error:", "Test FAILED", "Test OK")
    Exit Sub
Fail:
    PaneStatus frm, "Test error: " & Err.Description
End Sub

Public Sub Pane_RunPrompt(frm As Object)
    On Error GoTo Fail
    Dim q As String: q = Trim$(frm.txtPrompt.Text)
    If q = "" Then PaneStatus frm, "Enter a prompt first": Exit Sub
    ApplyPaneConfig frm
    PaneStatus frm, "Running prompt..."
    Dim sys As String
    sys = Trim$(frm.txtSystem.Text)
    If sys = "" Then sys = "You are a helpful assistant embedded in Excel."
    Dim r As String
    r = ChatComplete(sys, q, CurrentProvider, CurrentModel)
    PaneOut frm, ">> " & q & vbCrLf & r
    PaneRefreshUsage frm
    PaneStatus frm, "Done"
    Exit Sub
Fail:
    PaneStatus frm, "Prompt error: " & Err.Description
End Sub

Public Sub Pane_RunAgent(frm As Object)
    On Error GoTo Fail
    Dim q As String: q = Trim$(frm.txtPrompt.Text)
    If q = "" Then PaneStatus frm, "Enter an instruction first": Exit Sub
    ApplyPaneConfig frm
    Dim auto As Boolean: auto = (frm.chkAuto.Value = True)
    PaneStatus frm, "Running agent" & IIf(auto, " (auto-apply)...", " (approve prompt may appear)...")
    Dim r As String
    r = RunAgentLoop(q, 8, CurrentProvider, CurrentModel, auto)
    PaneOut frm, "[AGENT] " & q & vbCrLf & r
    PaneRefreshUsage frm
    PaneStatus frm, "Agent finished"
    Exit Sub
Fail:
    PaneStatus frm, "Agent error: " & Err.Description
End Sub

Public Sub Pane_ResetUsage(frm As Object)
    On Error Resume Next
    ResetUsage
    PaneRefreshUsage frm
    PaneStatus frm, "Usage reset"
End Sub

' ---- helpers ----------------------------------------------------------------

Private Sub ApplyPaneConfig(frm As Object)
    On Error Resume Next
    CurrentProvider = LCase$(Trim$(CStr(frm.cboProvider.Value)))
    CurrentModel = Trim$(frm.cboModel.Text)
    SetKeyForProvider CurrentProvider, frm.txtKey.Text
    SetBaseUrlForProvider CurrentProvider, Trim$(frm.txtBaseUrl.Text)
    ' MCP server (live; used by the agent's tool set this session)
    modMcp.McpUrl = Trim$(frm.txtMcp.Text)
    ' Image generation (BFL) key
    If Trim$(frm.txtImgKey.Text) <> "" Then SetProviderKey "bfl", Trim$(frm.txtImgKey.Text)
End Sub

Private Sub SetKeyForProvider(ByVal provider As String, ByVal key As String)
    Select Case LCase$(Trim$(provider))
        Case "openai": OPENAI_API_KEY = key
        Case "mistral": MISTRAL_API_KEY = key
        Case "nebius": NEBIUS_API_KEY = key
        Case "scaleway": SCALEWAY_API_KEY = key
        Case "openrouter": OPENROUTER_API_KEY = key
        Case "ollama"        ' local, no key
        Case Else: SetProviderKey provider, key
    End Select
End Sub

' Only the named providers keep an overridable base-URL variable; the extra
' OpenAI-compatible providers use fixed literals in GetBaseURL, so an override
' there is ignored (leave the field informational).
Private Sub SetBaseUrlForProvider(ByVal provider As String, ByVal url As String)
    If url = "" Then Exit Sub
    Select Case LCase$(Trim$(provider))
        Case "openai": OPENAI_URL = url
        Case "mistral": MISTRAL_URL = url
        Case "nebius": NEBIUS_URL = url
        Case "scaleway": SCALEWAY_URL = url
        Case "openrouter": OPENROUTER_URL = url
        Case "ollama": OLLAMA_BASE_URL = url
    End Select
End Sub

Private Sub PaneRefreshUsage(frm As Object)
    On Error Resume Next
    frm.lblUsage.Caption = UsageSummary()
End Sub

Private Sub PaneStatus(frm As Object, ByVal s As String)
    On Error Resume Next
    frm.lblStatus.Caption = s
    DoEvents
End Sub

Private Sub PaneOut(frm As Object, ByVal s As String)
    On Error Resume Next
    Dim cur As String
    cur = frm.txtOutput.Text
    If cur <> "" Then cur = cur & vbCrLf & String$(28, "-") & vbCrLf
    frm.txtOutput.Text = cur & s
    frm.txtOutput.SelStart = Len(frm.txtOutput.Text)
End Sub

' Curated fallback catalog per provider (used when the live /models call fails or
' returns nothing - e.g. offline, or a key-required provider with no key set).
Public Function DefaultModels(ByVal provider As String) As String
    Dim m As String
    Select Case LCase$(Trim$(provider))
        Case "ollama":      m = "llama3.2|qwen2.5|mistral|phi4|gemma2"
        Case "openai":      m = "gpt-4o|gpt-4o-mini|o3-mini|gpt-4.1|gpt-4.1-mini"
        Case "mistral":     m = "mistral-large-latest|mistral-small-latest|ministral-8b-latest|open-mistral-nemo"
        Case "nebius":      m = "meta-llama/Llama-3.3-70B-Instruct|Qwen/Qwen3-235B-A22B-Instruct-2507|Qwen/Qwen3-32B|google/gemma-3-27b-it|openai/gpt-oss-120b|deepseek-ai/DeepSeek-V3"
        Case "scaleway":    m = "llama-3.3-70b-instruct|qwen3-235b-a22b-instruct-2507|mistral-small-3.2-24b-instruct-2506|gpt-oss-120b|gemma-3-27b-it"
        Case "openrouter":  m = "openai/gpt-4o-mini|anthropic/claude-3.5-sonnet|meta-llama/llama-3.3-70b-instruct|google/gemini-flash-1.5"
        Case "groq":        m = "llama-3.3-70b-versatile|llama-3.1-8b-instant|mixtral-8x7b-32768"
        Case "together":    m = "meta-llama/Llama-3.3-70B-Instruct-Turbo|Qwen/Qwen2.5-72B-Instruct-Turbo"
        Case "cerebras":    m = "llama-3.3-70b|llama3.1-8b"
        Case "gemini":      m = "gemini-2.0-flash|gemini-1.5-pro|gemini-1.5-flash"
        Case "cohere":      m = "command-r-plus|command-r"
        Case "huggingface": m = "meta-llama/Llama-3.3-70B-Instruct|Qwen/Qwen2.5-72B-Instruct"
        Case "requesty":    m = "openai/gpt-4o-mini|anthropic/claude-3.5-sonnet"
        Case Else:          m = ""
    End Select
    DefaultModels = Replace(m, "|", vbCrLf)
End Function
