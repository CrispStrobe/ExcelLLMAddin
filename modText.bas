Attribute VB_Name = "modText"
' modText: UTF-8 encoding/decoding and text helpers.
'
' Root-cause fix for the recurring "umlaut" / mojibake bugs. The previous
' approach read UTF-8 bytes into a VBA string as if they were Latin-1 and then
' find-replaced specific broken sequences (see the old FixEncoding). That only
' patched the handful of characters someone happened to notice.
'
' Here we decode/encode UTF-8 <-> VBA UTF-16 correctly and platform-independently
' (no ADODB.Stream, which is unavailable on Mac). These functions are pure and
' fully unit-tested in modTests.
Option Explicit

' Decode a UTF-8 byte array into a proper VBA (UTF-16) String.
' Malformed sequences are replaced with U+FFFD rather than throwing.
Public Function Utf8BytesToString(ByRef Bytes() As Byte) As String
    On Error GoTo Fail

    Dim lo As Long, hi As Long
    If Not TryArrayBounds(Bytes, lo, hi) Then
        Utf8BytesToString = ""
        Exit Function
    End If

    Dim Buffer As String
    Dim outPos As Long
    Dim i As Long
    Dim b0 As Long, cp As Long

    ' Output code units are never more numerous than input bytes.
    Buffer = Space$(hi - lo + 1)
    outPos = 0
    i = lo

    Do While i <= hi
        b0 = Bytes(i)
        If b0 < &H80 Then
            cp = b0
            i = i + 1
        ElseIf (b0 And &HE0) = &HC0 Then
            If i + 1 > hi Then
                cp = &HFFFD&: i = i + 1
            Else
                cp = ((b0 And &H1F) * &H40&) + (Bytes(i + 1) And &H3F)
                i = i + 2
            End If
        ElseIf (b0 And &HF0) = &HE0 Then
            If i + 2 > hi Then
                cp = &HFFFD&: i = i + 1
            Else
                cp = ((b0 And &HF) * &H1000&) + ((Bytes(i + 1) And &H3F) * &H40&) + (Bytes(i + 2) And &H3F)
                i = i + 3
            End If
        ElseIf (b0 And &HF8) = &HF0 Then
            If i + 3 > hi Then
                cp = &HFFFD&: i = i + 1
            Else
                cp = ((b0 And &H7) * &H40000) + ((Bytes(i + 1) And &H3F) * &H1000&) + _
                     ((Bytes(i + 2) And &H3F) * &H40&) + (Bytes(i + 3) And &H3F)
                i = i + 4
            End If
        Else
            cp = &HFFFD&: i = i + 1
        End If

        If cp <= &HFFFF& Then
            outPos = outPos + 1
            Mid$(Buffer, outPos, 1) = ChrWSafe(cp)
        Else
            ' Astral plane -> UTF-16 surrogate pair.
            cp = cp - &H10000
            outPos = outPos + 1
            Mid$(Buffer, outPos, 1) = ChrWSafe(&HD800& Or (cp \ &H400&))
            outPos = outPos + 1
            Mid$(Buffer, outPos, 1) = ChrWSafe(&HDC00& Or (cp And &H3FF))
        End If
    Loop

    Utf8BytesToString = Left$(Buffer, outPos)
    Exit Function

Fail:
    Utf8BytesToString = ""
End Function

' Encode a VBA (UTF-16) String into a UTF-8 byte array.
Public Function StringToUtf8Bytes(ByVal s As String) As Byte()
    Dim outBytes() As Byte
    Dim n As Long, i As Long, outLen As Long
    Dim code As Long, lo As Long, cp As Long

    n = Len(s)
    If n = 0 Then
        StringToUtf8Bytes = EmptyBytes()
        Exit Function
    End If

    ' Worst case 4 bytes per UTF-16 code unit.
    ReDim outBytes(0 To n * 4 - 1)
    outLen = 0
    i = 1

    Do While i <= n
        code = AscW(Mid$(s, i, 1)) And &HFFFF&
        If code >= &HD800& And code <= &HDBFF& And i < n Then
            lo = AscW(Mid$(s, i + 1, 1)) And &HFFFF&
            If lo >= &HDC00& And lo <= &HDFFF& Then
                cp = &H10000 + ((code - &HD800&) * &H400&) + (lo - &HDC00&)
                i = i + 2
            Else
                ' Unpaired high surrogate -> U+FFFD (mirror the decoder).
                cp = &HFFFD&: i = i + 1
            End If
        ElseIf code >= &HD800& And code <= &HDFFF& Then
            ' Lone surrogate (unpaired high at end, or an isolated low) -> U+FFFD.
            cp = &HFFFD&: i = i + 1
        Else
            cp = code: i = i + 1
        End If

        If cp < &H80 Then
            outBytes(outLen) = cp: outLen = outLen + 1
        ElseIf cp < &H800 Then
            outBytes(outLen) = &HC0 Or (cp \ &H40&): outLen = outLen + 1
            outBytes(outLen) = &H80 Or (cp And &H3F): outLen = outLen + 1
        ElseIf cp < &H10000 Then
            outBytes(outLen) = &HE0 Or (cp \ &H1000&): outLen = outLen + 1
            outBytes(outLen) = &H80 Or ((cp \ &H40&) And &H3F): outLen = outLen + 1
            outBytes(outLen) = &H80 Or (cp And &H3F): outLen = outLen + 1
        Else
            outBytes(outLen) = &HF0 Or (cp \ &H40000): outLen = outLen + 1
            outBytes(outLen) = &H80 Or ((cp \ &H1000&) And &H3F): outLen = outLen + 1
            outBytes(outLen) = &H80 Or ((cp \ &H40&) And &H3F): outLen = outLen + 1
            outBytes(outLen) = &H80 Or (cp And &H3F): outLen = outLen + 1
        End If
    Loop

    ReDim Preserve outBytes(0 To outLen - 1)
    StringToUtf8Bytes = outBytes
End Function

' Round-trip helper: UTF-8 encode a string and return it as another string
' whose bytes are the UTF-8 representation. Handy for tests and diagnostics.
Public Function Utf8RoundTrip(ByVal s As String) As String
    Utf8RoundTrip = Utf8BytesToString(StringToUtf8Bytes(s))
End Function

' Write a string to a file as UTF-8 bytes (no BOM). Used by the curl transport
' for request/config files and by the test harness for JUnit output.
Public Sub WriteUtf8File(ByVal path As String, ByVal content As String)
    Dim fileNum As Integer
    Dim b() As Byte
    Dim lo As Long, hi As Long
    b = StringToUtf8Bytes(content)
    fileNum = FreeFile
    Open path For Binary Access Write As #fileNum
    If TryArrayBounds(b, lo, hi) Then Put #fileNum, , b
    Close #fileNum
End Sub

' Read a whole file as UTF-8 into a proper VBA string.
Public Function ReadUtf8File(ByVal path As String) As String
    On Error GoTo Fail
    Dim fileNum As Integer, n As Long
    Dim b() As Byte
    fileNum = FreeFile
    Open path For Binary Access Read As #fileNum
    n = LOF(fileNum)
    If n = 0 Then
        Close #fileNum
        ReadUtf8File = ""
        Exit Function
    End If
    ReDim b(0 To n - 1)
    Get #fileNum, , b
    Close #fileNum
    ReadUtf8File = Utf8BytesToString(b)
    Exit Function
Fail:
    On Error Resume Next
    If fileNum > 0 Then Close #fileNum
    ReadUtf8File = ""
End Function

' ChrW that tolerates code units above &H7FFF (VBA ChrW takes a signed Integer).
Private Function ChrWSafe(ByVal codeUnit As Long) As String
    If codeUnit > &H7FFF& Then
        ChrWSafe = ChrW$(codeUnit - &H10000)
    Else
        ChrWSafe = ChrW$(codeUnit)
    End If
End Function

' Safely obtain the bounds of a possibly-uninitialized byte array.
Private Function TryArrayBounds(ByRef Bytes() As Byte, ByRef lo As Long, ByRef hi As Long) As Boolean
    On Error GoTo NotAllocated
    lo = LBound(Bytes)
    hi = UBound(Bytes)
    TryArrayBounds = (hi >= lo)
    Exit Function
NotAllocated:
    TryArrayBounds = False
End Function

Private Function EmptyBytes() As Byte()
    Dim b() As Byte
    EmptyBytes = b
End Function
