Attribute VB_Name = "modTests"
' modTests: self-contained test harness for the LLM add-in.
'
' Run interactively: put the cursor in RunAllTests and press F5, or call
' RunAllTests from the Immediate window (?RunAllTests(True)).
' Run headless (CI): Application.Run "RunAllTests", False  -> returns fail count
' and writes a JUnit XML report next to the workbook / in the temp dir.
'
' The tests need NO network and NO real provider: MockHttpClient is injected via
' modHttp.SetHttpClientForTest, so the full build-request -> parse-response
' pipeline runs deterministically. Expected non-ASCII values are built with
' ChrW so this source file stays pure ASCII and imports identically everywhere.
Option Explicit

Private mPass As Long
Private mFail As Long
Private mLog As String
Private mXml As String

' ---- entry point ------------------------------------------------------------

Public Function RunAllTests(Optional ByVal showUI As Boolean = True) As Long
    mPass = 0: mFail = 0: mLog = "": mXml = ""

    Dim savedProvider As String, savedModel As String, savedKey As String
    savedProvider = CurrentProvider
    savedModel = CurrentModel
    savedKey = OPENAI_API_KEY

    On Error GoTo Cleanup

    ' Deterministic config so EnsureConfig never reads a real file or the network.
    Call InitializeDefaults
    CurrentProvider = "ollama"
    CurrentModel = "test-model"
    OPENAI_API_KEY = "sk-test"

    ' --- pure UTF-8 layer ---
    Test_Utf8_Ascii
    Test_Utf8_Umlauts
    Test_Utf8_DecodeTwoByte
    Test_Utf8_Astral
    Test_Utf8_Empty
    Test_Utf8_LoneSurrogate

    ' --- request building ---
    Test_BuildBody_RoundTrip
    Test_Endpoints

    ' --- response parsing via prompt() ---
    Test_Prompt_Ollama
    Test_Prompt_OpenAI
    Test_Prompt_UnicodeEscape
    Test_Prompt_EmbeddedEscapes
    Test_Prompt_ErrorObject
    Test_Prompt_ErrorString
    Test_Prompt_TransportError

    ' --- model listing ---
    Test_ListModels_Ollama
    Test_ListModels_OpenAI

    ' --- caching ---
    Test_Cache_HitAvoidsSecondCall
    Test_Cache_DifferentInputsMiss
    Test_Cache_ErrorsNotCached

    ' --- task functions (modTasks) ---
    Test_Task_Helpers
    Test_Task_Classify
    Test_Task_Extract
    Test_Task_List
    Test_Task_Fields
    Test_Task_TolerantArray
    Test_Task_Cosine
    Test_Task_Embed

    ' --- agent (modAgent) ---
    Test_Agent_Tools
    Test_Agent_2DArray
    Test_Agent_HexBool
    Test_Agent_ChatWithTools

    ' --- MCP client (modMcp) ---
    Test_Mcp_ParseRpc
    Test_Mcp_ListTools
    Test_Mcp_CallTool

    ' --- validation guards ---
    Test_Prompt_MissingApiKey

Cleanup:
    modHttp.ResetHttpClient
    CurrentProvider = savedProvider
    CurrentModel = savedModel
    OPENAI_API_KEY = savedKey

    Dim summary As String
    summary = "Tests run: " & (mPass + mFail) & "   PASS: " & mPass & "   FAIL: " & mFail
    Debug.Print String(60, "=")
    Debug.Print mLog
    Debug.Print summary
    Debug.Print String(60, "=")

    WriteJUnit summary

    If showUI Then
        MsgBox summary, IIf(mFail = 0, vbInformation, vbExclamation), "ExcelLLM Test Results"
    End If

    RunAllTests = mFail
End Function

' ---- UTF-8 unit tests -------------------------------------------------------

Private Sub Test_Utf8_Ascii()
    AssertEqual "utf8/ascii round-trip", "Hello, world! 123", Utf8RoundTrip("Hello, world! 123")
End Sub

Private Sub Test_Utf8_Umlauts()
    ' "Gruesse Cafe" with real umlauts, built ASCII-safe.
    Dim s As String
    s = "Gr" & ChrW$(252) & ChrW$(223) & "e Caf" & ChrW$(233) & " " & _
        ChrW$(196) & ChrW$(214) & ChrW$(220)   ' AE OE UE
    AssertEqual "utf8/umlaut round-trip", s, Utf8RoundTrip(s)
End Sub

Private Sub Test_Utf8_DecodeTwoByte()
    ' 0xC3 0xA4 is UTF-8 for U+00E4 (a-umlaut).
    Dim b(0 To 1) As Byte
    b(0) = &HC3: b(1) = &HA4
    AssertEqual "utf8/decode 2-byte", ChrW$(&HE4), Utf8BytesToString(b)
End Sub

Private Sub Test_Utf8_Astral()
    ' 0xF0 0x9F 0x98 0x80 is UTF-8 for U+1F600 (grinning face).
    Dim b(0 To 3) As Byte
    b(0) = &HF0: b(1) = &H9F: b(2) = &H98: b(3) = &H80
    Dim emoji As String
    emoji = Utf8BytesToString(b)
    AssertEqual "utf8/astral decodes to surrogate pair (len 2)", "2", CStr(Len(emoji))
    AssertEqual "utf8/astral round-trip", emoji, Utf8RoundTrip(emoji)
End Sub

Private Sub Test_Utf8_Empty()
    AssertEqual "utf8/empty round-trip", "", Utf8RoundTrip("")
End Sub

Private Sub Test_Utf8_LoneSurrogate()
    ' A lone (unpaired) high surrogate must encode to U+FFFD, symmetric with the
    ' decoder. ChrW$(-10240) is code unit U+D800; ChrW$(-3) is U+FFFD.
    Dim loneHigh As String
    loneHigh = ChrW$(-10240)
    AssertEqual "utf8/lone surrogate -> U+FFFD", ChrW$(-3), Utf8RoundTrip(loneHigh)
End Sub

' ---- request building -------------------------------------------------------

Private Sub Test_BuildBody_RoundTrip()
    ' A prompt full of characters that naive escaping breaks: quotes, newline,
    ' umlaut, and an astral emoji. It must survive serialize -> parse intact.
    Dim b(0 To 3) As Byte
    b(0) = &HF0: b(1) = &H9F: b(2) = &H98: b(3) = &H80
    Dim emoji As String: emoji = Utf8BytesToString(b)

    Dim userText As String
    userText = "He said " & Chr$(34) & "hi" & Chr$(34) & vbLf & _
               "Gr" & ChrW$(252) & ChrW$(223) & "e " & emoji & " C:\path"

    Dim mock As MockHttpClient
    Set mock = InstallMock("{""message"":{""content"":""ok""}}")
    prompt userText, "ollama", "test-model"

    Dim root As Object, messages As Object, userMsg As Object
    Set root = JsonConverter.ParseJson(mock.LastBody)
    Set messages = root("messages")
    Set userMsg = messages(2)

    AssertEqual "body/model preserved", "test-model", CStr(root("model"))
    AssertEqual "body/stream flag for ollama", "False", CStr(root("stream"))
    AssertEqual "body/user content round-trips through JSON", userText, CStr(userMsg("content"))
End Sub

Private Sub Test_Endpoints()
    Dim mock As MockHttpClient

    Set mock = InstallMock("{""message"":{""content"":""ok""}}")
    prompt "x", "ollama", "test-model"
    AssertTrue "endpoint/ollama uses /api/chat", (InStr(mock.LastUrl, "/api/chat") > 0)

    Set mock = InstallMock("{""choices"":[{""message"":{""content"":""ok""}}]}")
    prompt "x", "openai", "gpt-4o-mini"
    AssertTrue "endpoint/openai uses /chat/completions", (InStr(mock.LastUrl, "/chat/completions") > 0)
    AssertEqual "endpoint/openai sends bearer key", "sk-test", mock.LastApiKey
End Sub

' ---- response parsing -------------------------------------------------------

Private Sub Test_Prompt_Ollama()
    InstallMock "{""message"":{""content"":""Hello Excel""},""done"":true}"
    AssertEqual "prompt/ollama basic", "Hello Excel", prompt("hi", "ollama", "test-model")
End Sub

Private Sub Test_Prompt_OpenAI()
    InstallMock "{""choices"":[{""message"":{""role"":""assistant"",""content"":""Hi there""}}]}"
    AssertEqual "prompt/openai basic", "Hi there", prompt("hi", "openai", "gpt-4o-mini")
End Sub

Private Sub Test_Prompt_UnicodeEscape()
    ' \u escapes must be decoded by the JSON parser (the old scanner never did).
    ' Fixture is pure ASCII on purpose: U+00E9 (e-acute) and U+00EF (i-diaeresis).
    InstallMock "{""message"":{""content"":""Caf\u00e9 na\u00efve""}}"
    Dim expected As String
    expected = "Caf" & ChrW$(233) & " na" & ChrW$(239) & "ve"
    AssertEqual "prompt/unicode escape decoded", expected, prompt("x", "ollama", "test-model")
End Sub

Private Sub Test_Prompt_EmbeddedEscapes()
    ' Embedded escaped quotes and newline inside the content string.
    InstallMock "{""message"":{""content"":""Line1\nLine2 \""q\""""}}"
    Dim expected As String
    expected = "Line1" & vbLf & "Line2 " & Chr$(34) & "q" & Chr$(34)
    AssertEqual "prompt/embedded escapes", expected, prompt("x", "ollama", "test-model")
End Sub

Private Sub Test_Prompt_ErrorObject()
    InstallMock "{""error"":{""message"":""invalid api key"",""type"":""auth""}}"
    AssertEqual "prompt/error object", "Error: invalid api key", prompt("x", "openai", "gpt-4o-mini")
End Sub

Private Sub Test_Prompt_ErrorString()
    InstallMock "{""error"":""model not found""}"
    AssertEqual "prompt/error string", "Error: model not found", prompt("x", "ollama", "test-model")
End Sub

Private Sub Test_Prompt_TransportError()
    ' Transport-layer failure surfaces verbatim.
    InstallMock "Error: Timeout - no response from curl"
    AssertEqual "prompt/transport error passthrough", _
        "Error: Timeout - no response from curl", prompt("x", "ollama", "test-model")
End Sub

' ---- model listing ----------------------------------------------------------

Private Sub Test_ListModels_Ollama()
    InstallMock "{""models"":[{""name"":""llama3.2""},{""name"":""mistral""}]}"
    AssertEqual "list/ollama", "llama3.2" & vbCrLf & "mistral", LIST_MODELS("ollama")
End Sub

Private Sub Test_ListModels_OpenAI()
    InstallMock "{""object"":""list"",""data"":[{""id"":""gpt-4o""},{""id"":""gpt-4o-mini""}]}"
    AssertEqual "list/openai", "gpt-4o" & vbCrLf & "gpt-4o-mini", LIST_MODELS("openai")
End Sub

' ---- validation guards ------------------------------------------------------

Private Sub Test_Prompt_MissingApiKey()
    Dim savedKey As String
    savedKey = OPENAI_API_KEY
    OPENAI_API_KEY = ""
    ' No mock needed: it must fail before any HTTP call.
    modHttp.ResetHttpClient
    AssertEqual "guard/missing api key", "Error: No API key for openai", _
        prompt("x", "openai", "gpt-4o-mini")
    OPENAI_API_KEY = savedKey
End Sub

' ---- caching ----------------------------------------------------------------

Private Sub Test_Cache_HitAvoidsSecondCall()
    Dim mock As MockHttpClient
    Set mock = InstallMock("{""message"":{""content"":""cached hi""}}")
    Dim r1 As String, r2 As String
    r1 = prompt("same", "ollama", "test-model")
    r2 = prompt("same", "ollama", "test-model")
    AssertEqual "cache/first result", "cached hi", r1
    AssertEqual "cache/second result matches", "cached hi", r2
    AssertEqual "cache/only one HTTP call", "1", CStr(mock.CallCount)
End Sub

Private Sub Test_Cache_DifferentInputsMiss()
    Dim mock As MockHttpClient
    Set mock = InstallMock("{""message"":{""content"":""x""}}")
    prompt "aaa", "ollama", "test-model"
    prompt "bbb", "ollama", "test-model"
    AssertEqual "cache/distinct prompts both call", "2", CStr(mock.CallCount)
End Sub

Private Sub Test_Cache_ErrorsNotCached()
    Dim mock As MockHttpClient
    Set mock = InstallMock("{""error"":""boom""}")
    prompt "err", "ollama", "test-model"
    prompt "err", "ollama", "test-model"
    AssertEqual "cache/errors retried not cached", "2", CStr(mock.CallCount)
End Sub

' ---- framework --------------------------------------------------------------

Private Sub Test_Task_Helpers()
    Dim c As Collection
    Set c = FlattenToStrings("a, b ,c")
    AssertEqual "task/flatten count", "3", CStr(c.Count)
    AssertEqual "task/flatten trims", "b", c(2)

    Dim cats As New Collection
    cats.Add "Positive": cats.Add "Negative"
    AssertEqual "task/match exact (case-insensitive)", "Positive", MatchCategory("positive", cats)
    AssertEqual "task/match contained", "Negative", MatchCategory("I'd say Negative", cats)

    Dim arr As Collection
    Set arr = ParseJsonStringArray("[""x"",""y"",""z""]")
    AssertEqual "task/parse array count", "3", CStr(arr.Count)
    AssertEqual "task/parse array item", "y", arr(2)
    Set arr = ParseJsonStringArray("```json" & vbLf & "[""a"",""b""]" & vbLf & "```")
    AssertEqual "task/parse array with fence", "a", arr(1)

    Dim items As Collection
    Set items = SplitLinesToItems("1. alpha" & vbLf & "2) beta" & vbLf & "- gamma")
    AssertEqual "task/split lines strips prefixes", "beta", items(2)
End Sub

Private Sub Test_Task_Classify()
    InstallMock "{""message"":{""content"":""Positive""}}"
    AssertEqual "task/classify", "Positive", CLASSIFY("great product", "Positive,Negative", "ollama", "test-model")
End Sub

Private Sub Test_Task_Extract()
    InstallMock "{""message"":{""content"":""   bob@x.com   ""}}"
    AssertEqual "task/extract trims", "bob@x.com", EXTRACT("mail bob@x.com", "the email", "ollama", "test-model")
End Sub

Private Sub Test_Task_List()
    InstallMock "{""message"":{""content"":""[\""red\"",\""green\"",\""blue\""]""}}"
    Dim r As Variant
    r = LIST("primary colors", 0, "ollama", "test-model")
    AssertEqual "task/list first", "red", CStr(r(1, 1))
    AssertEqual "task/list third", "blue", CStr(r(3, 1))
End Sub

Private Sub Test_Task_Fields()
    InstallMock "{""message"":{""content"":""[\""Bob\"",\""bob@x.com\""]""}}"
    Dim r As Variant
    r = FIELDS("Bob bob@x.com", "name,email", "ollama", "test-model")
    AssertEqual "task/fields name", "Bob", CStr(r(1, 1))
    AssertEqual "task/fields email", "bob@x.com", CStr(r(1, 2))
End Sub

Private Sub Test_Task_TolerantArray()
    ' Strict JSON still parses.
    Dim c As Collection
    Set c = ParseJsonStringArray("[""x"",""y""]")
    AssertTrue "task/tolerant strict not nothing", Not (c Is Nothing)
    AssertEqual "task/tolerant strict count", "2", CStr(c.Count)

    ' A trailing comma (common from small models) is repaired, not rejected.
    Dim t As Collection
    Set t = ParseJsonStringArray("[""a"",""b"",]")
    AssertTrue "task/tolerant trailing-comma not nothing", Not (t Is Nothing)
    AssertEqual "task/tolerant trailing-comma count", "2", CStr(t.Count)
    AssertEqual "task/tolerant trailing-comma first", "a", CStr(t(1))
    AssertEqual "task/tolerant trailing-comma second", "b", CStr(t(2))

    ' Trailing comma with whitespace before the bracket, too.
    Dim w As Collection
    Set w = ParseJsonStringArray("[""a"", ""b"" , ]")
    AssertTrue "task/tolerant ws-comma not nothing", Not (w Is Nothing)
    AssertEqual "task/tolerant ws-comma count", "2", CStr(w.Count)
End Sub

Private Sub Test_Task_Cosine()
    Dim a As New Collection, b As New Collection
    a.Add 1#: a.Add 2#: a.Add 3#
    b.Add 1#: b.Add 2#: b.Add 3#
    AssertEqual "task/cosine identical", "1", CStr(Round(Cosine(a, b), 4))

    Dim c As New Collection, d As New Collection
    c.Add 1#: c.Add 0#
    d.Add 0#: d.Add 1#
    AssertEqual "task/cosine orthogonal", "0", CStr(Round(Cosine(c, d), 4))
End Sub

Private Sub Test_Task_Embed()
    InstallMock "{""data"":[{""embedding"":[1.0,2.0,3.0]}]}"
    Dim v As Object
    Set v = EmbedVector("hello", "some-model", "openai")
    If v Is Nothing Then
        AssertEqual "task/embed parses vector", "3", "(nothing)"
    Else
        AssertEqual "task/embed vector length", "3", CStr(v.Count)
        AssertEqual "task/embed vector value", "2", CStr(v(2))
    End If
End Sub

Private Sub Test_Agent_Tools()
    Dim tools As Collection
    Set tools = GetAgentTools()
    AssertEqual "agent/tools present", "True", CStr(tools.Count >= 5)
    AssertEqual "agent/write_range is a write tool", "True", CStr(IsWriteTool("write_range"))
    AssertEqual "agent/read_range is not a write tool", "False", CStr(IsWriteTool("read_range"))
End Sub

Private Sub Test_Agent_2DArray()
    Dim outer As New Collection, r1 As New Collection, r2 As New Collection
    r1.Add 1: r1.Add 2
    r2.Add 3: r2.Add 4
    outer.Add r1: outer.Add r2
    Dim a As Variant
    a = CollectionTo2DArray(outer)
    AssertEqual "agent/2darray dims", "2x2", CStr(UBound(a, 1)) & "x" & CStr(UBound(a, 2))
    AssertEqual "agent/2darray value", "4", CStr(a(2, 2))
End Sub

Private Sub Test_Agent_HexBool()
    AssertEqual "agent/hex color", CStr(RGB(255, 235, 156)), CStr(HexToColor("#FFEB9C"))
    AssertEqual "agent/bool from string", "True", CStr(ToBool("true"))
    AssertEqual "agent/bool from boolean", "True", CStr(ToBool(True))
    AssertEqual "agent/bool false", "False", CStr(ToBool("no"))
End Sub

Private Sub Test_Agent_ChatWithTools()
    InstallMock "{""choices"":[{""message"":{""role"":""assistant"",""content"":null,""tool_calls"":[{""id"":""c1"",""type"":""function"",""function"":{""name"":""write_range"",""arguments"":""{}""}}]}}]}"
    Dim msgs As New Collection
    msgs.Add MsgDict("system", "s")
    msgs.Add MsgDict("user", "u")
    Dim assistant As Object
    Set assistant = ChatWithTools(msgs, GetAgentTools(), "openai", "gpt-4o-mini")
    If assistant Is Nothing Then
        AssertEqual "agent/ChatWithTools parses tool_calls", "write_range", "(nothing)"
    Else
        AssertEqual "agent/ChatWithTools parses tool_calls", "write_range", CStr(assistant("tool_calls")(1)("function")("name"))
    End If
End Sub

Private Sub Test_Mcp_ParseRpc()
    Dim m As Object
    Set m = ParseRpcResult("{""jsonrpc"":""2.0"",""id"":5,""result"":{""x"":1}}", 5)
    AssertEqual "mcp/parse json result", "1", CStr(m("result")("x"))
    Set m = ParseRpcResult("event: message" & vbLf & "data: {""jsonrpc"":""2.0"",""id"":6,""result"":{""y"":2}}" & vbLf, 6)
    AssertEqual "mcp/parse SSE result", "2", CStr(m("result")("y"))
End Sub

Private Sub Test_Mcp_ListTools()
    InstallMock "{""jsonrpc"":""2.0"",""id"":2,""result"":{""tools"":[{""name"":""search"",""description"":""web search"",""inputSchema"":{""type"":""object"",""properties"":{""q"":{""type"":""string""}}}}]}}"
    Dim t As Collection
    Set t = McpListTools("https://mcp.example/rpc", "")
    If t Is Nothing Then
        AssertEqual "mcp/list tools", "search", "(nothing)"
    Else
        AssertEqual "mcp/list tools count", "1", CStr(t.Count)
        AssertEqual "mcp/list tool name", "search", CStr(t(1)("function")("name"))
    End If
End Sub

Private Sub Test_Mcp_CallTool()
    InstallMock "{""jsonrpc"":""2.0"",""id"":3,""result"":{""content"":[{""type"":""text"",""text"":""result text""}]}}"
    AssertEqual "mcp/call tool text", "result text", McpCallTool("https://mcp.example/rpc", "search", New Dictionary, "")
End Sub

Private Function InstallMock(ByVal response As String) As MockHttpClient
    ' Fresh cache per test so cross-test prompts don't collide on cache keys.
    ClearLLMCache True
    Dim mock As MockHttpClient
    Set mock = New MockHttpClient
    mock.NextResponse = response
    modHttp.SetHttpClientForTest mock
    Set InstallMock = mock
End Function

Private Sub AssertEqual(ByVal testName As String, ByVal expected As String, ByVal actual As String)
    If StrComp(expected, actual, vbBinaryCompare) = 0 Then
        RecordResult testName, True, ""
    Else
        RecordResult testName, False, "expected=[" & expected & "] actual=[" & actual & "]"
    End If
End Sub

Private Sub AssertTrue(ByVal testName As String, ByVal condition As Boolean)
    RecordResult testName, condition, "condition was False"
End Sub

Private Sub RecordResult(ByVal testName As String, ByVal ok As Boolean, ByVal message As String)
    If ok Then
        mPass = mPass + 1
        mLog = mLog & "PASS  " & testName & vbCrLf
    Else
        mFail = mFail + 1
        mLog = mLog & "FAIL  " & testName & vbCrLf & "        " & message & vbCrLf
    End If

    mXml = mXml & "  <testcase name=""" & XmlEsc(testName) & """>"
    If Not ok Then mXml = mXml & "<failure message=""" & XmlEsc(message) & """></failure>"
    mXml = mXml & "</testcase>" & vbLf
End Sub

Private Sub WriteJUnit(ByVal summary As String)
    On Error Resume Next
    Dim path As String, xml As String
    path = ReportPath()
    xml = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbLf & _
          "<testsuite name=""ExcelLLM"" tests=""" & (mPass + mFail) & _
          """ failures=""" & mFail & """>" & vbLf & _
          mXml & "</testsuite>" & vbLf
    modText.WriteUtf8File path, xml
    Debug.Print "JUnit report: " & path & " (" & summary & ")"
End Sub

Private Function ReportPath() As String
    Dim d As String
#If Mac Then
    d = Environ$("TMPDIR")
    If d = "" Then d = "/tmp/"
    If Right$(d, 1) <> "/" Then d = d & "/"
#Else
    d = Environ$("TEMP")
    If d = "" Then d = Environ$("TMP")
    If Right$(d, 1) <> "\" Then d = d & "\"
#End If
    ReportPath = d & "excelllm_junit.xml"
End Function

Private Function XmlEsc(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, Chr$(34), "&quot;")
    XmlEsc = s
End Function
