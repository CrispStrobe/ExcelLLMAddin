Attribute VB_Name = "modMcp"
' modMcp: a best-effort MCP (Model Context Protocol) client over HTTP (JSON-RPC
' 2.0), ported from officejs/src/mcp.ts. Lets the VBA agent use tools from a
' remote MCP server in addition to the built-in Excel tools.
'
' Caveats vs the JS client: IHttpClient is tailored for LLM calls, so we can't
' read the response's Mcp-Session-Id header or set a custom Accept header. This
' therefore works with stateless / lenient JSON-responding MCP servers. An
' optional bearer token is sent via the Authorization header. Configure with the
' SetMcpServer macro (session-scoped).
Option Explicit

Public McpUrl As String
Public McpToken As String

' Set the MCP server for this session (not persisted).
Public Sub SetMcpServer()
    McpUrl = InputBox("Remote MCP server URL (blank to disable):", "MCP Server", McpUrl)
    If McpUrl <> "" Then
        McpToken = InputBox("Optional bearer token (blank if none):", "MCP Server", McpToken)
    End If
    MsgBox IIf(McpUrl = "", "MCP disabled.", "MCP server set: " & McpUrl), vbInformation
End Sub

' Connect (initialize + tools/list) and return the tools as agent tool defs, or
' Nothing on failure.
Public Function McpListTools(url As String, Optional token As String = "") As Collection
    On Error GoTo Fail

    Dim initParams As New Dictionary
    initParams.Add "protocolVersion", "2024-11-05"
    initParams.Add "capabilities", New Dictionary
    Dim ci As New Dictionary
    ci.Add "name", "excel-llm-addin"
    ci.Add "version", "1.0"
    initParams.Add "clientInfo", ci

    ' Best-effort: proceed to tools/list even if initialize returns nothing.
    Dim ignored As Object
    Set ignored = McpRpc(url, "initialize", initParams, 1, token)

    Dim listResult As Object
    Set listResult = McpRpc(url, "tools/list", New Dictionary, 2, token)
    If listResult Is Nothing Then Set McpListTools = Nothing: Exit Function
    If Not listResult.Exists("tools") Then Set McpListTools = New Collection: Exit Function

    Dim tools As Object
    Set tools = listResult("tools")

    Dim out As New Collection, i As Long
    For i = 1 To tools.Count
        Dim tool As Object
        Set tool = tools(i)
        Dim d As New Dictionary
        d.Add "type", "function"
        Dim f As New Dictionary
        f.Add "name", tool("name")
        f.Add "description", IIf(tool.Exists("description"), CStr(tool("description")), "")
        If tool.Exists("inputSchema") Then
            f.Add "parameters", tool("inputSchema")
        Else
            Dim emptySchema As New Dictionary
            emptySchema.Add "type", "object"
            emptySchema.Add "properties", New Dictionary
            f.Add "parameters", emptySchema
        End If
        d.Add "function", f
        out.Add d
    Next i
    Set McpListTools = out
    Exit Function
Fail:
    Set McpListTools = Nothing
End Function

' Call a remote MCP tool; returns its text content, or "Error: ...".
Public Function McpCallTool(url As String, name As String, args As Object, Optional token As String = "") As String
    On Error GoTo Fail
    Dim params As New Dictionary
    params.Add "name", name
    If args Is Nothing Then Set args = New Dictionary
    params.Add "arguments", args

    Dim result As Object
    Set result = McpRpc(url, "tools/call", params, 3, token)
    If result Is Nothing Then McpCallTool = "Error: MCP call failed": Exit Function

    If result.Exists("content") Then
        Dim content As Object
        Set content = result("content")
        Dim s As String, i As Long
        For i = 1 To content.Count
            If content(i).Exists("text") Then s = s & IIf(s <> "", vbLf, "") & CStr(content(i)("text"))
        Next i
        McpCallTool = s
    Else
        McpCallTool = JsonConverter.ConvertToJson(result)
    End If
    Exit Function
Fail:
    McpCallTool = "Error: " & Err.Description
End Function

' One JSON-RPC call; returns the "result" object or Nothing (on transport/error).
Private Function McpRpc(url As String, method As String, params As Object, id As Long, token As String) As Object
    On Error GoTo Fail
    Dim req As New Dictionary
    req.Add "jsonrpc", "2.0"
    req.Add "id", id
    req.Add "method", method
    If Not params Is Nothing Then req.Add "params", params

    Dim client As IHttpClient
    Set client = modHttp.CreateHttpClient()

    ' token via the apiKey param -> Authorization: Bearer <token>; provider "" -> no extras.
    Dim response As String
    response = client.PostJson(url, JsonConverter.ConvertToJson(req), token, "")
    If Left$(response, 6) = "Error:" Then Set McpRpc = Nothing: Exit Function

    Dim msg As Object
    Set msg = ParseRpcResult(response, id)
    If msg Is Nothing Then Set McpRpc = Nothing: Exit Function
    If msg.Exists("error") Then Set McpRpc = Nothing: Exit Function
    If msg.Exists("result") Then Set McpRpc = msg("result") Else Set McpRpc = Nothing
    Exit Function
Fail:
    Set McpRpc = Nothing
End Function

' Extract the JSON-RPC message from a plain-JSON or SSE response body.
Public Function ParseRpcResult(response As String, id As Long) As Object
    On Error GoTo Fail
    Dim s As String
    s = Trim(response)

    If Left$(s, 1) = "{" Then
        Set ParseRpcResult = JsonConverter.ParseJson(s)
        Exit Function
    ElseIf Left$(s, 1) = "[" Then
        Dim arr As Object, i As Long
        Set arr = JsonConverter.ParseJson(s)
        For i = 1 To arr.Count
            If arr(i).Exists("id") Then
                If arr(i)("id") = id Then Set ParseRpcResult = arr(i): Exit Function
            End If
        Next i
        For i = 1 To arr.Count
            If arr(i).Exists("result") Or arr(i).Exists("error") Then Set ParseRpcResult = arr(i): Exit Function
        Next i
        Set ParseRpcResult = Nothing
        Exit Function
    Else
        ' SSE: scan data: lines for a JSON-RPC message.
        Dim lines() As String, ln As String, payload As String, k As Long
        lines = Split(Replace(s, vbCr, ""), vbLf)
        For k = LBound(lines) To UBound(lines)
            ln = Trim(lines(k))
            If Left$(ln, 5) = "data:" Then
                payload = Trim(Mid$(ln, 6))
                If payload <> "" And payload <> "[DONE]" Then
                    Dim m As Object
                    On Error Resume Next
                    Set m = JsonConverter.ParseJson(payload)
                    On Error GoTo Fail
                    If Not m Is Nothing Then
                        If m.Exists("result") Or m.Exists("error") Then Set ParseRpcResult = m: Exit Function
                    End If
                End If
            End If
        Next k
        Set ParseRpcResult = Nothing
    End If
    Exit Function
Fail:
    Set ParseRpcResult = Nothing
End Function
