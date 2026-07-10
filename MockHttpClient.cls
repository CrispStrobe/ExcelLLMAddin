VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MockHttpClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' MockHttpClient: test double implementing IHttpClient.
'
' Preload a canned response and inject via modHttp.SetHttpClientForTest. Then
' prompt()/LIST_MODELS() run with zero network and zero filesystem, so the whole
' request-build -> parse -> unescape pipeline is unit-testable and deterministic.
'
' It also records the last request (URL and body) so tests can assert that the
' correct endpoint and JSON payload were constructed per provider.
Option Explicit

Implements IHttpClient

Public NextResponse As String       ' the body PostJson/GetJson will return
Public NextGetResponse As String    ' optional separate response for GET

Public LastUrl As String
Public LastBody As String
Public LastApiKey As String
Public LastProvider As String
Public CallCount As Long

Private Function IHttpClient_PostJson(ByVal url As String, ByVal body As String, _
                                      ByVal apiKey As String, ByVal provider As String) As String
    LastUrl = url
    LastBody = body
    LastApiKey = apiKey
    LastProvider = provider
    CallCount = CallCount + 1
    IHttpClient_PostJson = NextResponse
End Function

Private Function IHttpClient_GetJson(ByVal url As String, _
                                     ByVal apiKey As String, ByVal provider As String) As String
    LastUrl = url
    LastApiKey = apiKey
    LastProvider = provider
    CallCount = CallCount + 1
    If NextGetResponse <> "" Then
        IHttpClient_GetJson = NextGetResponse
    Else
        IHttpClient_GetJson = NextResponse
    End If
End Function
