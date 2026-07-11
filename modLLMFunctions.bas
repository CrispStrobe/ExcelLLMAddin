Attribute VB_Name = "modLLMFunctions"
' modLLMFunctions: worksheet functions (PROMPT, LIST_MODELS, LLM_CONFIG) and the
' request/response pipeline.
'
' This module no longer talks to curl or the filesystem directly. It builds
' requests with a real JSON serializer (JsonConverter/ConvertToJson), sends them
' through an injected IHttpClient (see modHttp), and parses responses with a real
' JSON parser (JsonConverter/ParseJson). The old hand-rolled string scanners,
' EscapeJSON, and the FixEncoding mojibake table are gone: escaping is correct by
' construction and UTF-8 is handled at the transport boundary (modText/CurlClient
' /WinHttpClient).
Option Explicit

' Neutral, spreadsheet-appropriate system prompt. UTF-8 now works end to end, so
' we no longer need to beg the model to avoid non-ASCII characters.
Private Const SYSTEM_PROMPT As String = _
    "You are a helpful assistant embedded in a spreadsheet. " & _
    "Answer concisely and return plain text suitable for a cell unless asked otherwise."

' Session response cache, keyed by provider|model|prompt. Because prompt() is no
' longer Application.Volatile, a cell only recalculates when its inputs change;
' this cache additionally collapses identical prompts across the whole sheet to a
' single API call. Only successful results are cached (errors retry). Clear it
' with ClearLLMCache to force fresh answers.
Private mCache As Object

' Detect platform (diagnostics only; not exposed as a worksheet function).
Private Function IsMac() As Boolean
    #If Mac Then
        IsMac = True
    #Else
        IsMac = False
    #End If
End Function

' =PROMPT(text, [provider], [model])
' NOTE: intentionally NOT Application.Volatile. Volatile made every prompt cell
' re-call the API on every unrelated recalc (cost + latency + rate limits). A
' normal function still recalculates when its input cells change, which is what
' users actually want. Identical (provider, model, prompt) inputs are additionally
' served from a session cache (see mCache / ClearLLMCache).
Public Function prompt(promptText As String, Optional provider As String = "", Optional model As String = "") As String
    prompt = ChatComplete(SYSTEM_PROMPT, promptText, provider, model)
End Function

' Core chat completion: system + user message -> assistant text. Shared by PROMPT
' and every task function (CLASSIFY, EXTRACT, TRANSLATE, ...). Handles config,
' model resolution, caching, request building and response parsing. Public so the
' task modules can reuse it (it is not meant to be used directly in a cell).
Public Function ChatComplete(systemMsg As String, userMsg As String, Optional provider As String = "", Optional model As String = "") As String
    On Error GoTo ErrorHandler

    EnsureConfig

    If provider = "" Then provider = CurrentProvider
    If model = "" Then model = CurrentModel
    provider = LCase(Trim(provider))

    Dim baseURL As String, apiKey As String, endpoint As String
    Dim jsonBody As String, response As String

    baseURL = GetBaseURL(provider)
    apiKey = GetAPIKey(provider)

    If baseURL = "" Then
        ChatComplete = "Error: Invalid provider '" & provider & "'"
        Exit Function
    End If

    If apiKey = "" And provider <> "ollama" Then
        ChatComplete = "Error: No API key for " & provider
        Exit Function
    End If

    If model = "" Then
        model = ResolveDefaultModel(provider)
        If Left$(model, 6) = "Error:" Then
            ChatComplete = model
            Exit Function
        End If
    End If

    ' Cache lookup -- the key includes the system message so different task prompts
    ' (classify vs translate vs ...) don't collide.
    Dim cacheKey As String, cached As String
    cacheKey = provider & "|" & model & "|" & systemMsg & "|" & userMsg
    If TryGetCache(cacheKey, cached) Then
        ChatComplete = cached
        Exit Function
    End If

    Select Case provider
        Case "ollama"
            endpoint = baseURL & "/api/chat"
        Case Else
            endpoint = baseURL & "/chat/completions"
    End Select

    jsonBody = BuildChatBody(model, systemMsg, userMsg, provider)

    Dim client As IHttpClient
    Set client = modHttp.CreateHttpClient()
    response = client.PostJson(endpoint, jsonBody, apiKey, provider)

    If Left$(response, 6) = "Error:" Then
        ChatComplete = response
        Exit Function
    End If

    ChatComplete = ExtractChatContent(response, provider)

    ' Only cache genuine successes so transient failures can retry.
    If Left$(ChatComplete, 6) <> "Error:" Then StoreCache cacheKey, ChatComplete
    Exit Function

ErrorHandler:
    ChatComplete = "Error: " & Err.Number & " - " & Err.Description
End Function

' Embedding vector for SIMILARITY. Returns a Collection of Doubles, or Nothing on
' error. Openai-style (/embeddings) and Ollama (/api/embeddings) response shapes.
Public Function EmbedVector(text As String, embModel As String, Optional provider As String = "") As Object
    On Error GoTo Fail

    EnsureConfig
    If provider = "" Then provider = CurrentProvider
    provider = LCase(Trim(provider))
    If embModel = "" Then Set EmbedVector = Nothing: Exit Function

    Dim baseURL As String, apiKey As String, endpoint As String, response As String
    baseURL = GetBaseURL(provider)
    apiKey = GetAPIKey(provider)
    If baseURL = "" Then Set EmbedVector = Nothing: Exit Function
    If apiKey = "" And provider <> "ollama" Then Set EmbedVector = Nothing: Exit Function

    Dim body As Object
    Set body = New Dictionary
    body.Add "model", embModel
    If provider = "ollama" Then
        endpoint = baseURL & "/api/embeddings"
        body.Add "prompt", text
    Else
        endpoint = baseURL & "/embeddings"
        body.Add "input", text
    End If

    Dim client As IHttpClient
    Set client = modHttp.CreateHttpClient()
    response = client.PostJson(endpoint, JsonConverter.ConvertToJson(body), apiKey, provider)
    If Left$(response, 6) = "Error:" Then Set EmbedVector = Nothing: Exit Function

    Dim root As Object, vec As Object
    Set root = JsonConverter.ParseJson(response)
    If root.Exists("data") Then
        Set vec = root("data")(1)("embedding")          ' OpenAI style
    ElseIf root.Exists("embedding") Then
        Set vec = root("embedding")                      ' Ollama style
    ElseIf root.Exists("embeddings") Then
        Set vec = root("embeddings")(1)
    Else
        Set EmbedVector = Nothing: Exit Function
    End If

    Set EmbedVector = vec
    Exit Function
Fail:
    Set EmbedVector = Nothing
End Function

' =LIST_MODELS([provider]) -> newline-separated model ids.
Public Function LIST_MODELS(Optional provider As String = "") As String
    On Error GoTo ErrorHandler

    EnsureConfig
    If provider = "" Then provider = CurrentProvider
    provider = LCase(Trim(provider))

    Dim baseURL As String, apiKey As String, endpoint As String
    Dim response As String, models As String

    baseURL = GetBaseURL(provider)
    apiKey = GetAPIKey(provider)

    If baseURL = "" Then
        LIST_MODELS = "Error: Invalid provider"
        Exit Function
    End If
    If apiKey = "" And provider <> "ollama" Then
        LIST_MODELS = "Error: No API key for " & provider
        Exit Function
    End If

    Select Case provider
        Case "ollama": endpoint = baseURL & "/api/tags"
        Case Else: endpoint = baseURL & "/models"
    End Select

    Dim client As IHttpClient
    Set client = modHttp.CreateHttpClient()
    response = client.GetJson(endpoint, apiKey, provider)

    If Left$(response, 6) = "Error:" Then
        LIST_MODELS = response
        Exit Function
    End If

    models = ParseModelList(response, provider)
    If Left$(models, 6) <> "Error:" Then
        LIST_MODELS = Replace(models, "|", vbCrLf)
    Else
        LIST_MODELS = models
    End If
    Exit Function

ErrorHandler:
    LIST_MODELS = "Error: " & Err.Description
End Function

' =LLM_CONFIG() -> current provider/model summary.
Public Function LLM_CONFIG() As String
    On Error Resume Next
    If CurrentProvider = "" Or CurrentModel = "" Then Call LoadConfig

    If CurrentProvider = "" Then
        LLM_CONFIG = "Status: Not Configured (Run ShowSettings)"
    Else
        LLM_CONFIG = "Provider: " & CurrentProvider & " | Model: " & CurrentModel
    End If
End Function

' Connectivity probe used by the menu. Now platform-independent (WinHttp on
' Windows, curl on Mac) via the same client the UDFs use.
Public Function TestCurl() As String
    On Error GoTo ErrorHandler
    Dim client As IHttpClient
    Set client = modHttp.CreateHttpClient()
    TestCurl = Trim$(client.GetJson("http://localhost:11434/api/version", "", "ollama"))
    Exit Function
ErrorHandler:
    TestCurl = "Error: " & Err.Description
End Function

' ---- internal helpers -------------------------------------------------------

Private Sub EnsureConfig()
    If CurrentProvider = "" Or OLLAMA_BASE_URL = "" Then Call LoadConfig
End Sub

' ---- response cache ---------------------------------------------------------

' Empty the session response cache. Run this (Tools > Macro) to force =PROMPT()
' cells to re-query the model instead of returning cached answers.
' Note: cached cells won't refresh until they recalculate (e.g. Ctrl/Cmd+Alt+F9).
Public Sub ClearLLMCache(Optional ByVal silent As Boolean = False)
    Set mCache = Nothing
    If Not silent Then MsgBox "LLM response cache cleared.", vbInformation, "LLM Add-in"
End Sub

Private Function TryGetCache(ByVal key As String, ByRef value As String) As Boolean
    If mCache Is Nothing Then Exit Function
    If mCache.Exists(key) Then
        value = mCache(key)
        TryGetCache = True
    End If
End Function

Private Sub StoreCache(ByVal key As String, ByVal value As String)
    If mCache Is Nothing Then Set mCache = New Dictionary
    mCache(key) = value
End Sub

' Pick a usable model when the caller/config didn't specify one.
Private Function ResolveDefaultModel(ByVal provider As String) As String
    Dim available As String
    available = LIST_MODELS(provider)

    If Left$(available, 6) = "Error:" Then
        ResolveDefaultModel = "Error: No valid model available for " & provider
        Exit Function
    End If

    If CurrentModel <> "" And InStr(available, CurrentModel) > 0 Then
        ResolveDefaultModel = CurrentModel
    ElseIf available <> "" Then
        ResolveDefaultModel = Split(available, vbCrLf)(0)
    Else
        ResolveDefaultModel = "Error: No valid model available for " & provider
    End If
End Function

' Build a chat-completions request body with a real serializer so prompt text
' (quotes, backslashes, newlines, unicode) is always escaped correctly.
' Multimodal chat: ask a question about an image (OpenAI-compatible content array).
' Direct providers only; Ollama uses a different image format so it is rejected.
Public Function ChatCompleteVision(question As String, imageUrl As String, _
                                   Optional provider As String = "", Optional model As String = "") As String
    On Error GoTo ErrorHandler
    EnsureConfig
    If provider = "" Then provider = CurrentProvider
    If model = "" Then model = CurrentModel
    provider = LCase(Trim(provider))

    If provider = "ollama" Then
        ChatCompleteVision = "Error: VISION isn't supported for Ollama (different image format)"
        Exit Function
    End If

    Dim baseURL As String, apiKey As String
    baseURL = GetBaseURL(provider)
    apiKey = GetAPIKey(provider)
    If baseURL = "" Then ChatCompleteVision = "Error: Invalid provider '" & provider & "'": Exit Function
    If apiKey = "" Then ChatCompleteVision = "Error: No API key for " & provider: Exit Function

    ' Build the multimodal message: content = [ {type:text}, {type:image_url} ].
    Dim root As Object, messages As Object, um As Object, content As Object
    Dim txt As Object, img As Object, imgUrl As Object
    Set root = New Dictionary
    root.Add "model", model
    Set messages = New Collection

    Set um = New Dictionary
    um.Add "role", "user"
    Set content = New Collection
    Set txt = New Dictionary
    txt.Add "type", "text": txt.Add "text", question
    content.Add txt
    Set img = New Dictionary
    img.Add "type", "image_url"
    Set imgUrl = New Dictionary
    imgUrl.Add "url", imageUrl
    img.Add "image_url", imgUrl
    content.Add img
    um.Add "content", content
    messages.Add um
    root.Add "messages", messages

    Dim jsonBody As String
    jsonBody = JsonConverter.ConvertToJson(root)

    Dim client As IHttpClient
    Set client = modHttp.CreateHttpClient()
    Dim response As String
    response = client.PostJson(baseURL & "/chat/completions", jsonBody, apiKey, provider)
    If Left$(response, 6) = "Error:" Then ChatCompleteVision = response: Exit Function

    ChatCompleteVision = ExtractChatContent(response, provider)
    Exit Function
ErrorHandler:
    ChatCompleteVision = "Error: " & Err.Number & " - " & Err.Description
End Function

Private Function BuildChatBody(ByVal model As String, ByVal systemMsg As String, _
                               ByVal userMsg As String, ByVal provider As String) As String
    Dim root As Object, messages As Object, m As Object
    Set root = New Dictionary
    root.Add "model", model

    Set messages = New Collection
    Set m = New Dictionary
    m.Add "role", "system": m.Add "content", systemMsg
    messages.Add m
    Set m = New Dictionary
    m.Add "role", "user": m.Add "content", userMsg
    messages.Add m
    root.Add "messages", messages

    If provider = "ollama" Then root.Add "stream", False

    BuildChatBody = JsonConverter.ConvertToJson(root)
End Function

' Extract assistant text from a chat response, tolerant of provider shape:
'   Ollama:   {"message":{"content":"..."}}
'   OpenAI-*: {"choices":[{"message":{"content":"..."}}]}
Private Function ExtractChatContent(ByVal jsonText As String, ByVal provider As String) As String
    On Error GoTo Fail

    Dim root As Object
    Set root = JsonConverter.ParseJson(jsonText)

    If root.Exists("error") Then
        ExtractChatContent = "Error: " & ExtractErrorMessage(root("error"))
        Exit Function
    End If

    Dim msg As Object, choices As Object, c0 As Object

    ' OpenAI-compatible shape first (covers most providers, incl. Ollama's
    ' /v1 compat) then Ollama native.
    If root.Exists("choices") Then
        Set choices = root("choices")
        If choices.Count >= 1 Then
            Set c0 = choices(1)
            If c0.Exists("message") Then
                Set msg = c0("message")
                ExtractChatContent = CStr(msg("content"))
                Exit Function
            ElseIf c0.Exists("text") Then
                ExtractChatContent = CStr(c0("text"))
                Exit Function
            End If
        End If
    End If

    If root.Exists("message") Then
        Set msg = root("message")
        If msg.Exists("content") Then
            ExtractChatContent = CStr(msg("content"))
            Exit Function
        End If
    End If

    ExtractChatContent = "Error: No content in response: " & Left$(jsonText, 200)
    Exit Function

Fail:
    ExtractChatContent = "Error: Parse failed - " & Err.Description
End Function

Private Function ExtractErrorMessage(ByVal errVal As Variant) As String
    If IsObject(errVal) Then
        Dim d As Object
        Set d = errVal
        If d.Exists("message") Then
            ExtractErrorMessage = CStr(d("message"))
        Else
            ExtractErrorMessage = "Unknown error"
        End If
    Else
        ExtractErrorMessage = CStr(errVal)
    End If
End Function

' Parse a models listing into a pipe-delimited string.
'   Ollama /api/tags: {"models":[{"name":"..."}]}
'   OpenAI /models:   {"data":[{"id":"..."}]}
Private Function ParseModelList(ByVal jsonText As String, ByVal provider As String) As String
    On Error GoTo Fail

    Dim root As Object, arr As Object, item As Object
    Dim result As String, i As Long
    Set root = JsonConverter.ParseJson(jsonText)
    provider = LCase(Trim(provider))

    If provider = "ollama" Then
        If root.Exists("models") Then
            Set arr = root("models")
            For i = 1 To arr.Count
                Set item = arr(i)
                If item.Exists("name") Then result = AppendPipe(result, CStr(item("name")))
            Next i
        End If
    Else
        If root.Exists("data") Then
            Set arr = root("data")
            For i = 1 To arr.Count
                Set item = arr(i)
                If item.Exists("id") Then result = AppendPipe(result, CStr(item("id")))
            Next i
        End If
    End If

    If result = "" Then
        ParseModelList = "Error: No models found"
    Else
        ParseModelList = result
    End If
    Exit Function

Fail:
    ParseModelList = "Error: " & Err.Description
End Function

Private Function AppendPipe(ByVal acc As String, ByVal item As String) As String
    If acc = "" Then
        AppendPipe = item
    Else
        AppendPipe = acc & "|" & item
    End If
End Function
