Attribute VB_Name = "modHttp"
' modHttp: HTTP client factory + session helpers.
'
' CreateHttpClient() returns the right IHttpClient for the platform (WinHttp on
' Windows, curl on Mac). Tests call SetHttpClientForTest to swap in a mock, then
' ResetHttpClient to restore production behaviour.
Option Explicit

Private mTokenSeq As Long
Private mTestClient As IHttpClient

' Return the transport to use. Honors a test override if one is installed.
Public Function CreateHttpClient() As IHttpClient
    If Not mTestClient Is Nothing Then
        Set CreateHttpClient = mTestClient
        Exit Function
    End If

#If Mac Then
    Set CreateHttpClient = New CurlClient
#Else
    Set CreateHttpClient = New WinHttpClient
#End If
End Function

' Install a test double (mock/fake) as the transport for all subsequent calls.
Public Sub SetHttpClientForTest(ByVal client As IHttpClient)
    Set mTestClient = client
End Sub

' Remove any test override and return to the real platform transport.
Public Sub ResetHttpClient()
    Set mTestClient = Nothing
End Sub

' Session-unique token for temp filenames. Combines a wall-clock stamp with a
' monotonically increasing counter, so even multiple recalcs within the same
' second get distinct filenames.
Public Function UniqueToken() As String
    mTokenSeq = mTokenSeq + 1
    UniqueToken = Format$(Now, "yyyymmddhhnnss") & "_" & Format$(mTokenSeq, "000000")
End Function
