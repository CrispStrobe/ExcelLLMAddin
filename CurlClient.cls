VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CurlClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' CurlClient: Mac (and non-Windows fallback) IHttpClient implementation.
'
' Mac Excel has no in-process HTTPS object, so we shell out to curl. Compared to
' the original inline curl code this fixes three real defects:
'   1. UTF-8 correctness: the request body is written as UTF-8 bytes and the
'      response is read as UTF-8 bytes, both via modText. No more mojibake.
'   2. Secrets & injection: the URL, headers, API key and body go into files
'      (a curl --config (-K) file + a data-binary body file) instead of the
'      command line, so keys never appear in `ps` and no caller/prompt content is
'      interpolated into a shell string. The ONLY thing on the command line is
'      the config file path, which is system-derived (TempDir) plus an
'      alphanumeric session token -- no user-controlled content.
'   3. Temp-file collisions: filenames use a per-session unique token instead of
'      second-granularity timestamps, so concurrent recalcs can't clobber.
'
' On Mac, AppleScript's `do shell script` blocks until curl exits, so we read the
' output immediately with no polling. The Shell()-based branch (non-Mac fallback)
' still polls because Shell is asynchronous.
Option Explicit

Implements IHttpClient

Private Const MAX_TIME_POST As Long = 120   ' seconds
Private Const MAX_TIME_GET As Long = 30

Private Function IHttpClient_PostJson(ByVal url As String, ByVal body As String, _
                                      ByVal apiKey As String, ByVal provider As String) As String
    IHttpClient_PostJson = Execute("POST", url, body, apiKey, provider, MAX_TIME_POST)
End Function

Private Function IHttpClient_GetJson(ByVal url As String, _
                                     ByVal apiKey As String, ByVal provider As String) As String
    IHttpClient_GetJson = Execute("GET", url, "", apiKey, provider, MAX_TIME_GET)
End Function

Private Function Execute(ByVal method As String, ByVal url As String, ByVal body As String, _
                         ByVal apiKey As String, ByVal provider As String, ByVal maxTime As Long) As String
    On Error GoTo Fail

    Dim tmp As String, token As String
    Dim bodyFile As String, configFile As String, outputFile As String

    tmp = TempDir()
    token = modHttp.UniqueToken()
    bodyFile = tmp & "llm_body_" & token & ".json"
    configFile = tmp & "llm_cfg_" & token & ".conf"
    outputFile = tmp & "llm_out_" & token & ".txt"

    If method = "POST" Then modText.WriteUtf8File bodyFile, body
    modText.WriteUtf8File configFile, BuildCurlConfig(method, url, apiKey, provider, bodyFile, outputFile, maxTime)

    Dim ok As Boolean
    ok = RunCurl(configFile, outputFile)

    If Not ok Then
        Execute = "Error: Timeout - no response from curl"
    Else
        Execute = modText.ReadUtf8File(outputFile)
        If Execute = "" Then Execute = "Error: Empty response"
    End If

    CleanupFile bodyFile
    CleanupFile configFile
    CleanupFile outputFile
    Exit Function

Fail:
    Execute = "Error: " & Err.Number & " - " & Err.Description
    On Error Resume Next
    CleanupFile bodyFile
    CleanupFile configFile
    CleanupFile outputFile
End Function

' Build a curl config (-K) file. Keeps URL/key/headers off the command line.
Private Function BuildCurlConfig(ByVal method As String, ByVal url As String, ByVal apiKey As String, _
                                 ByVal provider As String, ByVal bodyFile As String, _
                                 ByVal outputFile As String, ByVal maxTime As Long) As String
    Dim c As String
    c = "url = """ & CfgEsc(url) & """" & vbLf
    c = c & "silent" & vbLf
    c = c & "show-error" & vbLf
    c = c & "max-time = " & maxTime & vbLf
    c = c & "output = """ & CfgEsc(outputFile) & """" & vbLf

    If method = "POST" Then
        c = c & "request = ""POST""" & vbLf
        c = c & "header = ""Content-Type: application/json""" & vbLf
        c = c & "data-binary = ""@" & CfgEsc(bodyFile) & """" & vbLf
    End If

    If apiKey <> "" Then
        If LCase$(provider) = "bfl" Then
            ' BFL (image generation) authenticates with x-key, not Bearer.
            c = c & "header = ""x-key: " & CfgEsc(apiKey) & """" & vbLf
        Else
            c = c & "header = ""Authorization: Bearer " & CfgEsc(apiKey) & """" & vbLf
        End If
    End If
    If LCase$(provider) = "openrouter" Then
        c = c & "header = ""HTTP-Referer: https://excel-addin""" & vbLf
        c = c & "header = ""X-Title: ExcelLLM Add-in""" & vbLf
    End If

    BuildCurlConfig = c
End Function

' Escape a value for inclusion inside a double-quoted curl config token.
Private Function CfgEsc(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    CfgEsc = s
End Function

' Returns True once the output file exists. On Mac the call is synchronous so
' this is immediate; on the Shell() fallback we poll until curl finishes writing.
Private Function RunCurl(ByVal configFile As String, ByVal outputFile As String) As Boolean
#If Mac Then
    Dim script As String
    script = "with timeout of 320 seconds" & vbCr & _
             "do shell script ""curl -K '" & configFile & "'""" & vbCr & _
             "end timeout"
    MacScript script
    RunCurl = (Dir(outputFile) <> "")
#Else
    Shell "cmd /c curl -K """ & configFile & """", vbHide
    RunCurl = WaitForFile(outputFile)
#End If
End Function

' Fallback-only: wait for an async Shell'd curl to create and finish the file.
Private Function WaitForFile(ByVal path As String) As Boolean
    Dim waitCount As Long, lastSize As Long, curSize As Long, stable As Long
    Do While Dir(path) = "" And waitCount < 60
        Application.Wait Now + TimeValue("00:00:01")
        waitCount = waitCount + 1
    Loop
    If Dir(path) = "" Then Exit Function

    lastSize = -1: stable = 0: waitCount = 0
    Do While stable < 3 And waitCount < 30
        On Error Resume Next
        curSize = FileLen(path)
        On Error GoTo 0
        If curSize = lastSize And curSize > 0 Then
            stable = stable + 1
        Else
            stable = 0: lastSize = curSize
        End If
        Application.Wait Now + TimeValue("00:00:01")
        waitCount = waitCount + 1
    Loop
    WaitForFile = True
End Function

Private Function TempDir() As String
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
    TempDir = d
End Function

Private Sub CleanupFile(ByVal path As String)
    On Error Resume Next
    If path <> "" Then
        If Dir(path) <> "" Then Kill path
    End If
End Sub
