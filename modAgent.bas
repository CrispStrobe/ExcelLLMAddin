Attribute VB_Name = "modAgent"
' modAgent: a tool-calling agent that edits the active worksheet, ported from the
' Office.js agent (officejs/src/core/agent.ts + excelTools.ts). Works offline with
' local Ollama. The model is given Excel "tools" (read/write ranges, formulas,
' formatting, sheets) and loops via function-calling until done. Mutating tools
' are queued and applied only after the user approves.
'
' Entry point: run the RunAgent macro (Tools > Macro), or wire a shortcut.
Option Explicit

Private Const AGENT_SYSTEM As String = _
    "You are an assistant operating on the user's Excel worksheet via tools. " & _
    "Read what you need, then make the requested changes in small, verifiable steps. " & _
    "Prefer writing formulas over hardcoded values when it fits. Address ranges use A1 " & _
    "notation (e.g. Sheet1!B2:B10). When the task is done, reply with a short summary " & _
    "and no further tool call."

' MCP tool names active this run, for routing tool calls to the MCP server.
Private mMcpNames As Object

' ---- entry point ------------------------------------------------------------

Public Sub RunAgent()
    Dim instruction As String
    instruction = InputBox( _
        "Describe the change for the agent to make on this sheet:" & vbLf & vbLf & _
        "e.g. In D1 put the sum of B2:B10, then bold anything over 100", _
        "LLM Agent")
    If Trim(instruction) = "" Then Exit Sub

    Application.StatusBar = "Agent working..."
    Dim result As String
    result = RunAgentLoop(instruction)
    Application.StatusBar = False

    MsgBox result, vbInformation, "LLM Agent"
End Sub

' Run the tool-calling loop. Returns a log + final summary. Write tools are queued
' and, at the end, applied only if the user confirms.
Public Function RunAgentLoop(instruction As String, Optional maxSteps As Long = 8, _
                             Optional provider As String = "", Optional model As String = "") As String
    On Error GoTo Fail

    Dim messages As New Collection
    messages.Add MsgDict("system", AGENT_SYSTEM)
    messages.Add MsgDict("user", instruction)

    Dim tools As Collection
    Set tools = GetAgentTools()

    ' Merge in remote MCP tools if a server is configured (see modMcp.SetMcpServer).
    Set mMcpNames = New Dictionary
    If modMcp.McpUrl <> "" Then
        Dim mcpTools As Collection
        Set mcpTools = modMcp.McpListTools(modMcp.McpUrl, modMcp.McpToken)
        If Not mcpTools Is Nothing Then
            Dim mi As Long
            For mi = 1 To mcpTools.Count
                tools.Add mcpTools(mi)
                mMcpNames(CStr(mcpTools(mi)("function")("name"))) = True
            Next mi
        End If
    End If

    Dim pending As New Collection
    Dim log As String
    Dim iStep As Long

    For iStep = 1 To maxSteps
        Dim assistant As Object
        Set assistant = ChatWithTools(messages, tools, provider, model)
        If assistant Is Nothing Then
            RunAgentLoop = log & vbLf & "Error: no/invalid response from the model."
            GoTo ApplyPending
        End If

        Dim toolCalls As Object
        Set toolCalls = Nothing
        If assistant.Exists("tool_calls") Then Set toolCalls = assistant("tool_calls")

        If toolCalls Is Nothing Then
            RunAgentLoop = log & vbLf & vbLf & FinalText(assistant)
            GoTo ApplyPending
        ElseIf toolCalls.Count = 0 Then
            RunAgentLoop = log & vbLf & vbLf & FinalText(assistant)
            GoTo ApplyPending
        End If

        messages.Add assistant

        Dim i As Long
        For i = 1 To toolCalls.Count
            Dim tc As Object
            Set tc = toolCalls(i)
            Dim fname As String
            fname = CStr(tc("function")("name"))

            ' OpenAI returns arguments as a JSON string; Ollama returns an object.
            Dim argsObj As Object
            If IsObject(tc("function")("arguments")) Then
                Set argsObj = tc("function")("arguments")
            Else
                Set argsObj = ParseArgs(CStr(tc("function")("arguments")))
            End If

            Dim result As String
            If IsWriteTool(fname) Then
                Dim pa As New Dictionary
                pa.Add "name", fname
                pa.Add "args", argsObj
                pending.Add pa
                result = "Queued " & fname & " for approval (not applied yet)."
            Else
                result = ExecTool(fname, argsObj)
            End If

            log = log & vbLf & "- " & fname & " -> " & Left$(result, 140)
            messages.Add ToolMsg(CStr(tc("id")), result)
        Next i
    Next iStep

    RunAgentLoop = log & vbLf & vbLf & "Stopped: reached the step limit."

ApplyPending:
    If pending.Count > 0 Then
        Dim summary As String, j As Long
        For j = 1 To pending.Count
            summary = summary & "- " & pending(j)("name") & "(" & Left$(ArgsPreview(pending(j)("args")), 80) & ")" & vbLf
        Next j
        If MsgBox("Apply these " & pending.Count & " change(s)?" & vbLf & vbLf & summary, _
                  vbYesNo + vbQuestion, "LLM Agent") = vbYes Then
            For j = 1 To pending.Count
                Dim r As String
                r = ExecTool(CStr(pending(j)("name")), pending(j)("args"))
                RunAgentLoop = RunAgentLoop & vbLf & "applied " & pending(j)("name") & " -> " & Left$(r, 80)
            Next j
        Else
            RunAgentLoop = RunAgentLoop & vbLf & "(changes not applied)"
        End If
    End If
    Exit Function

Fail:
    RunAgentLoop = "Error: " & Err.Description
End Function

' One turn: send messages + tools, return the assistant message object (with
' content and/or tool_calls), or Nothing on error.
Public Function ChatWithTools(messages As Collection, tools As Collection, _
                              Optional provider As String = "", Optional model As String = "") As Object
    On Error GoTo Fail

    EnsureConfig
    If provider = "" Then provider = CurrentProvider
    If model = "" Then model = CurrentModel
    provider = LCase(Trim(provider))

    Dim baseURL As String, apiKey As String, endpoint As String
    baseURL = GetBaseURL(provider)
    apiKey = GetAPIKey(provider)
    If baseURL = "" Then Set ChatWithTools = Nothing: Exit Function
    If apiKey = "" And provider <> "ollama" Then Set ChatWithTools = Nothing: Exit Function

    Dim root As New Dictionary
    root.Add "model", model
    root.Add "messages", messages
    root.Add "tools", tools
    root.Add "tool_choice", "auto"
    If provider = "ollama" Then root.Add "stream", False

    Select Case provider
        Case "ollama": endpoint = baseURL & "/api/chat"
        Case Else: endpoint = baseURL & "/chat/completions"
    End Select

    Dim client As IHttpClient
    Set client = modHttp.CreateHttpClient()
    Dim response As String
    response = client.PostJson(endpoint, JsonConverter.ConvertToJson(root), apiKey, provider)
    If Left$(response, 6) = "Error:" Then Set ChatWithTools = Nothing: Exit Function

    Dim resp As Object
    Set resp = JsonConverter.ParseJson(response)
    If resp.Exists("error") Then Set ChatWithTools = Nothing: Exit Function

    Set ChatWithTools = resp("choices")(1)("message")
    Exit Function
Fail:
    Set ChatWithTools = Nothing
End Function

' ---- tool schemas -----------------------------------------------------------

Public Function GetAgentTools() As Collection
    Dim t As New Collection
    Dim props As Dictionary

    Set props = New Dictionary
    props.Add "address", StrProp("A1-style range, e.g. Sheet1!B2:B10")
    t.Add ToolDef("read_range", "Read the values of a range in A1 notation.", ObjSchema(props, ReqArr("address")))

    t.Add ToolDef("get_selection", "Get the address and values of the current selection.", ObjSchema(New Dictionary, Nothing))
    t.Add ToolDef("list_sheets", "List the worksheet names.", ObjSchema(New Dictionary, Nothing))

    Set props = New Dictionary
    props.Add "address", StrProp("top-left cell/range")
    props.Add "values", ArrProp("2D array of values")
    t.Add ToolDef("write_range", "Write a 2D array of values starting at a top-left cell (resizes to fit).", ObjSchema(props, ReqArr2("address", "values")))

    Set props = New Dictionary
    props.Add "address", StrProp("range to fill")
    props.Add "formula", StrProp("e.g. =A2*B2")
    t.Add ToolDef("write_formula", "Fill a range with a formula (relative refs adjust per cell).", ObjSchema(props, ReqArr2("address", "formula")))

    Set props = New Dictionary
    props.Add "address", StrProp("range to format")
    props.Add "fill", StrProp("background hex, e.g. #FFEB9C")
    props.Add "fontColor", StrProp("font hex")
    props.Add "bold", BoolProp()
    props.Add "italic", BoolProp()
    props.Add "numberFormat", StrProp("e.g. 0.00")
    t.Add ToolDef("set_format", "Format a range: fill, font color, bold, italic, number format.", ObjSchema(props, ReqArr("address")))

    Set props = New Dictionary
    props.Add "name", StrProp("new sheet name")
    t.Add ToolDef("add_worksheet", "Add a new worksheet.", ObjSchema(props, ReqArr("name")))

    Set props = New Dictionary
    props.Add "address", StrProp("A1 range holding the chart data (include headers)")
    props.Add "chartType", StrProp("ColumnClustered, BarClustered, Line, Pie, XYScatter, or Area")
    props.Add "title", StrProp("optional chart title")
    t.Add ToolDef("create_chart", "Create a chart from a data range.", ObjSchema(props, ReqArr2("address", "chartType")))

    Set GetAgentTools = t
End Function

Public Function IsWriteTool(fname As String) As Boolean
    Select Case fname
        Case "write_range", "write_formula", "set_format", "add_worksheet", "create_chart": IsWriteTool = True
        Case Else: IsWriteTool = False
    End Select
End Function

' ---- tool execution (native Excel) ------------------------------------------

Private Function ExecTool(fname As String, args As Object) As String
    On Error GoTo Fail
    If Not mMcpNames Is Nothing Then
        If mMcpNames.Exists(fname) Then
            ExecTool = modMcp.McpCallTool(modMcp.McpUrl, fname, args, modMcp.McpToken)
            Exit Function
        End If
    End If
    Select Case fname
        Case "read_range": ExecTool = Tool_ReadRange(CStr(args("address")))
        Case "get_selection": ExecTool = Tool_GetSelection()
        Case "list_sheets": ExecTool = Tool_ListSheets()
        Case "write_range": ExecTool = Tool_WriteRange(CStr(args("address")), args("values"))
        Case "write_formula": ExecTool = Tool_WriteFormula(CStr(args("address")), CStr(args("formula")))
        Case "set_format": ExecTool = Tool_SetFormat(args)
        Case "add_worksheet": ExecTool = Tool_AddWorksheet(CStr(args("name")))
        Case "create_chart": ExecTool = Tool_CreateChart(CStr(args("address")), CStr(args("chartType")), IIf(args.Exists("title"), CStr(args("title")), ""))
        Case Else: ExecTool = "Unknown tool: " & fname
    End Select
    Exit Function
Fail:
    ExecTool = "Error: " & Err.Description
End Function

Private Function RangeFromAddress(addr As String) As Range
    Dim bang As Long
    bang = InStr(addr, "!")
    If bang > 0 Then
        Dim sn As String
        sn = Replace(Left$(addr, bang - 1), "'", "")
        Set RangeFromAddress = Application.ActiveWorkbook.Worksheets(sn).Range(Mid$(addr, bang + 1))
    Else
        Set RangeFromAddress = Application.ActiveSheet.Range(addr)
    End If
End Function

Private Function Tool_ReadRange(addr As String) As String
    Dim rng As Range
    Set rng = RangeFromAddress(addr)
    Tool_ReadRange = rng.Address(False, False) & " = " & RangeValuesToJson(rng)
End Function

Private Function Tool_GetSelection() As String
    If TypeName(Application.Selection) <> "Range" Then Tool_GetSelection = "No range selected.": Exit Function
    Dim rng As Range
    Set rng = Application.Selection
    Tool_GetSelection = "Selection " & rng.Address(False, False) & " = " & RangeValuesToJson(rng)
End Function

Private Function Tool_ListSheets() As String
    Dim s As String, ws As Worksheet
    For Each ws In Application.ActiveWorkbook.Worksheets
        s = s & IIf(s <> "", ", ", "") & ws.Name
    Next ws
    Tool_ListSheets = "Sheets: " & s
End Function

Private Function Tool_WriteRange(addr As String, values As Variant) As String
    Dim arr2 As Variant
    arr2 = CollectionTo2DArray(values)
    Dim rows As Long, cols As Long
    rows = UBound(arr2, 1)
    cols = UBound(arr2, 2)
    If rows = 0 Or cols = 0 Then Tool_WriteRange = "No values to write.": Exit Function
    Dim rng As Range
    Set rng = RangeFromAddress(addr).Resize(rows, cols)
    rng.value = arr2
    Tool_WriteRange = "Wrote " & rows & "x" & cols & " to " & rng.Address(False, False)
End Function

Private Function Tool_WriteFormula(addr As String, formula As String) As String
    Dim rng As Range
    Set rng = RangeFromAddress(addr)
    rng.formula = formula
    Tool_WriteFormula = "Set formula " & formula & " on " & rng.Address(False, False)
End Function

Private Function Tool_SetFormat(args As Object) As String
    Dim rng As Range
    Set rng = RangeFromAddress(CStr(args("address")))
    If args.Exists("fill") Then rng.Interior.Color = HexToColor(CStr(args("fill")))
    If args.Exists("fontColor") Then rng.Font.Color = HexToColor(CStr(args("fontColor")))
    If args.Exists("bold") Then rng.Font.Bold = ToBool(args("bold"))
    If args.Exists("italic") Then rng.Font.Italic = ToBool(args("italic"))
    If args.Exists("numberFormat") Then rng.NumberFormat = CStr(args("numberFormat"))
    Tool_SetFormat = "Formatted " & rng.Address(False, False)
End Function

Private Function Tool_AddWorksheet(wsName As String) As String
    Dim ws As Worksheet
    Set ws = Application.ActiveWorkbook.Worksheets.Add
    ws.Name = wsName
    Tool_AddWorksheet = "Added worksheet '" & ws.Name & "'"
End Function

Private Function Tool_CreateChart(addr As String, chartType As String, title As String) As String
    Dim rng As Range
    Set rng = RangeFromAddress(addr)
    Dim ws As Worksheet
    Set ws = rng.Worksheet
    Dim ch As ChartObject
    Set ch = ws.ChartObjects.Add(Left:=rng.Left + rng.Width + 10, Top:=rng.Top, Width:=360, Height:=240)
    ch.Chart.SetSourceData Source:=rng
    ch.Chart.chartType = ResolveVbaChartType(chartType)
    If title <> "" Then
        ch.Chart.HasTitle = True
        ch.Chart.ChartTitle.text = title
    End If
    Tool_CreateChart = "Created " & chartType & " chart from " & rng.Address(False, False)
End Function

' Map a loose chart-type name to an XlChartType. Column is the safe default.
Private Function ResolveVbaChartType(ByVal t As String) As Long
    Dim k As String, i As Long, ch As String
    k = LCase$(Trim$(t))
    Dim clean As String
    For i = 1 To Len(k)
        ch = Mid$(k, i, 1)
        If ch >= "a" And ch <= "z" Then clean = clean & ch
    Next i
    Select Case clean
        Case "bar", "barclustered": ResolveVbaChartType = xlBarClustered
        Case "line": ResolveVbaChartType = xlLine
        Case "pie": ResolveVbaChartType = xlPie
        Case "scatter", "xyscatter": ResolveVbaChartType = xlXYScatter
        Case "area": ResolveVbaChartType = xlArea
        Case Else: ResolveVbaChartType = xlColumnClustered
    End Select
End Function

' ---- helpers ----------------------------------------------------------------

Public Function MsgDict(role As String, content As String) As Object
    Dim d As New Dictionary
    d.Add "role", role
    d.Add "content", content
    Set MsgDict = d
End Function

Private Function ToolMsg(toolCallId As String, content As String) As Object
    Dim d As New Dictionary
    d.Add "role", "tool"
    d.Add "tool_call_id", toolCallId
    d.Add "content", content
    Set ToolMsg = d
End Function

Private Function FinalText(assistant As Object) As String
    If assistant.Exists("content") Then
        If Not IsNull(assistant("content")) Then FinalText = CStr(assistant("content"))
    End If
    If FinalText = "" Then FinalText = "(done)"
End Function

Private Function ParseArgs(jsonStr As String) As Object
    On Error GoTo Fail
    If Trim(jsonStr) = "" Then Set ParseArgs = New Dictionary: Exit Function
    Set ParseArgs = JsonConverter.ParseJson(jsonStr)
    Exit Function
Fail:
    Set ParseArgs = New Dictionary
End Function

Private Function ArgsPreview(args As Object) As String
    On Error Resume Next
    ArgsPreview = JsonConverter.ConvertToJson(args)
End Function

' Convert a JsonConverter array (Collection of rows/values) to a 2D array.
Public Function CollectionTo2DArray(v As Variant) As Variant
    Dim rows As Long, cols As Long, r As Long, c As Long
    Dim nested As Boolean
    rows = v.Count
    nested = False
    If rows >= 1 Then
        If IsObject(v(1)) Then nested = True
    End If

    If nested Then
        cols = v(1).Count
        Dim arr() As Variant
        ReDim arr(1 To rows, 1 To cols)
        For r = 1 To rows
            Dim rowc As Object
            Set rowc = v(r)
            For c = 1 To cols
                If c <= rowc.Count Then arr(r, c) = rowc(c) Else arr(r, c) = ""
            Next c
        Next r
        CollectionTo2DArray = arr
    Else
        cols = rows
        Dim a2() As Variant
        ReDim a2(1 To 1, 1 To cols)
        For c = 1 To cols
            a2(1, c) = v(c)
        Next c
        CollectionTo2DArray = a2
    End If
End Function

Private Function RangeValuesToJson(rng As Range) As String
    Dim v As Variant
    v = rng.value
    Dim outer As New Collection
    If IsArray(v) Then
        Dim r As Long, c As Long
        For r = 1 To UBound(v, 1)
            Dim row As New Collection
            For c = 1 To UBound(v, 2)
                row.Add v(r, c)
            Next c
            outer.Add row
        Next r
    Else
        Dim row2 As New Collection
        row2.Add v
        outer.Add row2
    End If
    RangeValuesToJson = JsonConverter.ConvertToJson(outer)
End Function

Public Function HexToColor(hex As String) As Long
    Dim h As String
    h = Replace(hex, "#", "")
    If Len(h) >= 6 Then
        HexToColor = RGB(CLng("&H" & Mid$(h, 1, 2)), CLng("&H" & Mid$(h, 3, 2)), CLng("&H" & Mid$(h, 5, 2)))
    Else
        HexToColor = 0
    End If
End Function

Public Function ToBool(v As Variant) As Boolean
    If VarType(v) = vbBoolean Then
        ToBool = v
    Else
        ToBool = (LCase$(Trim$(CStr(v))) = "true")
    End If
End Function

' -- tiny JSON-schema builders --

Private Function ToolDef(fname As String, desc As String, params As Object) As Object
    Dim d As New Dictionary
    d.Add "type", "function"
    Dim f As New Dictionary
    f.Add "name", fname
    f.Add "description", desc
    f.Add "parameters", params
    d.Add "function", f
    Set ToolDef = d
End Function

Private Function ObjSchema(props As Object, required As Object) As Object
    Dim s As New Dictionary
    s.Add "type", "object"
    s.Add "properties", props
    If Not required Is Nothing Then s.Add "required", required
    Set ObjSchema = s
End Function

Private Function StrProp(desc As String) As Object
    Dim p As New Dictionary
    p.Add "type", "string"
    If desc <> "" Then p.Add "description", desc
    Set StrProp = p
End Function

Private Function BoolProp() As Object
    Dim p As New Dictionary
    p.Add "type", "boolean"
    Set BoolProp = p
End Function

Private Function ArrProp(desc As String) As Object
    Dim p As New Dictionary
    p.Add "type", "array"
    If desc <> "" Then p.Add "description", desc
    Set ArrProp = p
End Function

Private Function ReqArr(a As String) As Collection
    Dim c As New Collection
    c.Add a
    Set ReqArr = c
End Function

Private Function ReqArr2(a As String, b As String) As Collection
    Dim c As New Collection
    c.Add a
    c.Add b
    Set ReqArr2 = c
End Function
