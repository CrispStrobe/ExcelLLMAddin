VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinHttpClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' WinHttpClient: Windows IHttpClient implementation.
'
' Uses WinHttp.WinHttpRequest.5.1 (late-bound, no project reference needed) for
' synchronous, in-process HTTPS. This removes the Mac-style curl + temp-file +
' polling machinery on Windows entirely: it is faster, has no shell-injection
' surface, and never leaks API keys onto a process command line.
'
' Bodies are sent as UTF-8 bytes and responses decoded from ResponseBody bytes,
' so encoding is correct and identical to the Mac path.
Option Explicit

Implements IHttpClient

Private Const RESOLVE_TIMEOUT As Long = 10000    ' ms
Private Const CONNECT_TIMEOUT As Long = 15000
Private Const SEND_TIMEOUT As Long = 20000
Private Const RECEIVE_TIMEOUT As Long = 120000

Private Function IHttpClient_PostJson(ByVal url As String, ByVal body As String, _
                                      ByVal apiKey As String, ByVal provider As String) As String
    On Error GoTo Fail

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts RESOLVE_TIMEOUT, CONNECT_TIMEOUT, SEND_TIMEOUT, RECEIVE_TIMEOUT

    http.Open "POST", url, False
    http.SetRequestHeader "Content-Type", "application/json"
    ApplyAuthHeader http, apiKey, provider
    ApplyProviderHeaders http, provider

    http.Send StringToUtf8Bytes(body)

    IHttpClient_PostJson = ReadResponse(http)
    Exit Function
Fail:
    IHttpClient_PostJson = "Error: " & Err.Number & " - " & Err.Description
End Function

Private Function IHttpClient_GetJson(ByVal url As String, _
                                     ByVal apiKey As String, ByVal provider As String) As String
    On Error GoTo Fail

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts RESOLVE_TIMEOUT, CONNECT_TIMEOUT, SEND_TIMEOUT, RECEIVE_TIMEOUT

    http.Open "GET", url, False
    ApplyAuthHeader http, apiKey, provider
    ApplyProviderHeaders http, provider

    http.Send

    IHttpClient_GetJson = ReadResponse(http)
    Exit Function
Fail:
    IHttpClient_GetJson = "Error: " & Err.Number & " - " & Err.Description
End Function

' Auth header: BFL (image generation) uses x-key; everything else uses Bearer.
Private Sub ApplyAuthHeader(ByVal http As Object, ByVal apiKey As String, ByVal provider As String)
    If apiKey = "" Then Exit Sub
    If LCase$(provider) = "bfl" Then
        http.SetRequestHeader "x-key", apiKey
    Else
        http.SetRequestHeader "Authorization", "Bearer " & apiKey
    End If
End Sub

Private Sub ApplyProviderHeaders(ByVal http As Object, ByVal provider As String)
    If LCase$(provider) = "openrouter" Then
        http.SetRequestHeader "HTTP-Referer", "https://excel-addin"
        http.SetRequestHeader "X-Title", "ExcelLLM Add-in"
    End If
End Sub

' Decode the response as UTF-8 from the raw bytes. We return the body even for
' HTTP error statuses so the caller can surface the provider's JSON error
' message; only transport-level failures come back as "Error:".
Private Function ReadResponse(ByVal http As Object) As String
    Dim bytes() As Byte
    bytes = http.ResponseBody
    ReadResponse = Utf8BytesToString(bytes)
End Function
