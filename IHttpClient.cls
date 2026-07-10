VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IHttpClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' IHttpClient: transport abstraction for LLM HTTP calls.
'
' This interface is the testability seam. Production code talks to an
' IHttpClient; the factory in modHttp picks WinHttpClient (Windows, in-process)
' or CurlClient (Mac / fallback). Tests inject MockHttpClient with canned
' responses, so prompt(), LIST_MODELS() and all parsing can be exercised with
' zero network and zero filesystem.
'
' Convention (preserved from the original code): both methods return the raw
' response body on success, or a string beginning with "Error:" on transport
' failure. HTTP status handling and JSON parsing happen in the caller.
Option Explicit

' POST a JSON body. apiKey (if non-empty) is sent as a Bearer token.
' provider allows implementations to add provider-specific headers
' (e.g. OpenRouter's HTTP-Referer).
Public Function PostJson(ByVal url As String, ByVal body As String, _
                         ByVal apiKey As String, ByVal provider As String) As String
End Function

' GET a URL (used for model listing).
Public Function GetJson(ByVal url As String, _
                        ByVal apiKey As String, ByVal provider As String) As String
End Function
