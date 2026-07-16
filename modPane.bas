Attribute VB_Name = "modPane"
' modPane: logic for the modeless "task pane" UserForm (frmLLMPane).
'
' Pure VBA => works on locked-down Office where the Office.js web add-in is
' blocked by policy. This is a *floating* modeless panel (a true docked Custom
' Task Pane needs a COM/VSTO add-in), positioned at the right edge of the Excel
' window. The form's event handlers are one-liners that delegate here, so all
' behaviour lives in this text module (source-controlled + editable).
'
' Control names expected on frmLLMPane:
'   cboProvider, cboModel, txtKey, txtPrompt, txtOutput (multiline),
'   lblStatus, btnLoadModels, btnSave, btnTest, btnRunPrompt, btnRunAgent
Option Explicit

Private Const PANE_MARGIN As Long = 12

' Providers offered in the pane (matches GetAPIKey / GetBaseURL coverage).
Public Function PaneProviders() As Variant
    PaneProviders = Array("ollama", "openai", "mistral", "nebius", "scaleway", _
                          "openrouter", "groq", "together", "cerebras", _
                          "gemini", "cohere", "huggingface", "requesty")
End Function

' Public entry point: show the pane, positioned at the right edge of Excel.
Public Sub ShowLLMPane()
    On Error Resume Next
    LoadConfig
    With frmLLMPane
        .StartUpPosition = 0
        .Left = Application.Left + Application.Width - .Width - PANE_MARGIN
        .Top = Application.Top + 90
        .Show vbModeless
    End With
End Sub

Public Sub Pane_Init(frm As Object)
    On Error Resume Next
    Dim p As Variant
    frm.cboProvider.Clear
    For Each p In PaneProviders()
        frm.cboProvider.AddItem p
    Next p
    LoadConfig
    frm.cboProvider.Value = IIf(CurrentProvider = "", "ollama", CurrentProvider)
    frm.cboModel.Text = CurrentModel
    frm.txtKey.Text = GetAPIKey(CStr(frm.cboProvider.Value))
    frm.txtOutput.Text = ""
    PaneStatus frm, "Ready - provider: " & frm.cboProvider.Value
End Sub

Public Sub Pane_ProviderChanged(frm As Object)
    On Error Resume Next
    frm.txtKey.Text = GetAPIKey(CStr(frm.cboProvider.Value))
    PaneStatus frm, "Provider: " & frm.cboProvider.Value & _
        IIf(LCase$(CStr(frm.cboProvider.Value)) = "ollama", " (local, no key)", "")
End Sub

Public Sub Pane_LoadModels(frm As Object)
    On Error GoTo Fail
    ApplyPaneConfig frm
    Dim prov As String: prov = LCase$(Trim$(CStr(frm.cboProvider.Value)))
    PaneStatus frm, "Loading models for " & prov & "..."
    ' Live first (works keyless for OpenRouter/Ollama and with-key for the rest);
    ' fall back to the curated catalog so the dropdown is never empty.
    Dim s As String, src As String
    s = LIST_MODELS(prov)
    If Left$(s, 6) = "Error:" Or Trim$(s) = "" Then
        s = DefaultModels(prov)
        src = "defaults"
    Else
        src = "live"
    End If
    frm.cboModel.Clear
    Dim arr() As String, i As Long, n As Long
    arr = Split(s, vbCrLf)
    For i = LBound(arr) To UBound(arr)
        If Trim$(arr(i)) <> "" Then frm.cboModel.AddItem Trim$(arr(i)): n = n + 1
    Next i
    If n = 0 Then
        PaneStatus frm, "No models found for " & prov
    Else
        PaneStatus frm, n & " models (" & src & ")"
    End If
    Exit Sub
Fail:
    PaneStatus frm, "Model load error: " & Err.Description
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
    Dim r As String
    r = ChatComplete("You are a helpful assistant embedded in Excel.", q, _
                     CurrentProvider, CurrentModel)
    PaneOut frm, ">> " & q & vbCrLf & r
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
    PaneStatus frm, "Running agent (approve prompt may appear)..."
    Dim r As String
    r = RunAgentLoop(q, 8, CurrentProvider, CurrentModel)
    PaneOut frm, "[AGENT] " & q & vbCrLf & r
    PaneStatus frm, "Agent finished"
    Exit Sub
Fail:
    PaneStatus frm, "Agent error: " & Err.Description
End Sub

' ---- helpers ----------------------------------------------------------------

Private Sub ApplyPaneConfig(frm As Object)
    CurrentProvider = LCase$(Trim$(CStr(frm.cboProvider.Value)))
    CurrentModel = Trim$(frm.cboModel.Text)
    SetKeyForProvider CurrentProvider, frm.txtKey.Text
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
