Attribute VB_Name = "modTasks"
' modTasks: higher-level worksheet functions built on modLLMFunctions.ChatComplete
' (and EmbedVector for SIMILARITY). Ports the Office.js task functions to VBA so
' the .xlam reaches feature parity:
'   CLASSIFY, EXTRACT, TRANSLATE, SUMMARIZE, SENTIMENT, ASK, LIST, FIELDS, MAP,
'   SIMILARITY.
' The logic mirrors the unit-tested TypeScript in officejs/src/core/tasks.ts.
Option Explicit

' =CLASSIFY(text, categories, [provider], [model]) -> one of the labels.
Public Function CLASSIFY(text As String, categories As Variant, Optional provider As String = "", Optional model As String = "") As String
    On Error GoTo Fail
    Dim cats As Collection
    Set cats = FlattenToStrings(categories)
    If cats.Count = 0 Then CLASSIFY = "Error: no categories provided": Exit Function

    Dim labels As String, i As Long
    For i = 1 To cats.Count
        labels = labels & IIf(i > 1, ", ", "") & cats(i)
    Next i

    Dim sys As String
    sys = "You are a precise text classifier. Respond with EXACTLY ONE of the " & _
          "following labels and nothing else: " & labels & "."

    Dim out As String
    out = ChatComplete(sys, "Classify this text:" & vbLf & text, provider, model)
    If Left$(out, 6) = "Error:" Then CLASSIFY = out: Exit Function
    CLASSIFY = MatchCategory(out, cats)
    Exit Function
Fail:
    CLASSIFY = "Error: " & Err.Description
End Function

' =EXTRACT(text, instruction, ...) -> extracted value.
Public Function EXTRACT(text As String, instruction As String, Optional provider As String = "", Optional model As String = "") As String
    Dim sys As String
    sys = "Extract the requested information from the text. Output only the extracted " & _
          "value as plain text -- no labels, quotes, or explanation. If it is not present, output nothing."
    EXTRACT = TrimResult(ChatComplete(sys, "From the following text, extract: " & instruction & _
             vbLf & vbLf & "Text:" & vbLf & text, provider, model))
End Function

' =TRANSLATE(text, targetLanguage, ...) -> translation.
Public Function TRANSLATE(text As String, targetLanguage As String, Optional provider As String = "", Optional model As String = "") As String
    Dim sys As String
    sys = "You are a translator. Output only the translation, with no notes or quotes."
    TRANSLATE = TrimResult(ChatComplete(sys, "Translate the following into " & targetLanguage & ":" & _
                vbLf & vbLf & text, provider, model))
End Function

' =SUMMARIZE(text, [maxWords], ...) -> summary.
Public Function SUMMARIZE(text As String, Optional maxWords As Long = 0, Optional provider As String = "", Optional model As String = "") As String
    Dim limit As String
    If maxWords > 0 Then limit = " in at most " & maxWords & " words"
    Dim sys As String
    sys = "You are a concise summarizer. Output only the summary."
    SUMMARIZE = TrimResult(ChatComplete(sys, "Summarize the following" & limit & ":" & vbLf & vbLf & text, provider, model))
End Function

' =SENTIMENT(text, ...) -> Positive / Neutral / Negative.
Public Function SENTIMENT(text As String, Optional provider As String = "", Optional model As String = "") As String
    SENTIMENT = CLASSIFY(text, "Positive,Neutral,Negative", provider, model)
End Function

' =ASK(question, context, ...) -> answer using the context range/text.
Public Function ASK(question As String, context As Variant, Optional provider As String = "", Optional model As String = "") As String
    Dim ctx As String
    ctx = FlattenToText(context)
    Dim sys As String
    sys = "Answer the question using only the provided context. If the answer is not " & _
          "in the context, say so briefly. Output plain text suitable for a cell."
    ASK = TrimResult(ChatComplete(sys, "Context:" & vbLf & ctx & vbLf & vbLf & "Question: " & question, provider, model))
End Function

' =LIST(prompt, [count], ...) -> spills a column of items.
Public Function LIST(promptText As String, Optional count As Long = 0, Optional provider As String = "", Optional model As String = "") As Variant
    On Error GoTo Fail
    Dim ask As String
    If count > 0 Then ask = promptText & vbLf & vbLf & "Return exactly " & count & " items." Else ask = promptText

    Dim sys As String
    sys = "Answer as a JSON array of short strings -- no commentary, no code fences."
    Dim raw As String
    raw = ChatComplete(sys, ask, provider, model)
    If Left$(raw, 6) = "Error:" Then LIST = raw: Exit Function

    Dim items As Collection
    Set items = ParseJsonStringArray(raw)
    If items Is Nothing Then Set items = SplitLinesToItems(raw)

    Dim n As Long
    n = items.Count
    If count > 0 And count < n Then n = count
    If n = 0 Then LIST = "(no items)": Exit Function

    Dim arr() As Variant, i As Long
    ReDim arr(1 To n, 1 To 1)
    For i = 1 To n
        arr(i, 1) = items(i)
    Next i
    LIST = arr
    Exit Function
Fail:
    LIST = "Error: " & Err.Description
End Function

' =FIELDS(text, fields, ...) -> spills a row of extracted values.
Public Function FIELDS(text As String, fields As Variant, Optional provider As String = "", Optional model As String = "") As Variant
    On Error GoTo Fail
    Dim fs As Collection
    Set fs = FlattenToStrings(fields)
    If fs.Count = 0 Then FIELDS = "Error: no fields provided": Exit Function

    Dim numbered As String, i As Long
    For i = 1 To fs.Count
        numbered = numbered & i & ". " & fs(i) & vbLf
    Next i

    Dim sys As String
    sys = "Extract the requested fields from the text. Return ONLY a JSON array of " & _
          "string values, one per field, in the given order. Use an empty string for " & _
          "any field not present. No commentary or code fences."
    Dim user As String
    user = "Fields:" & vbLf & numbered & vbLf & "Text:" & vbLf & text & vbLf & vbLf & _
           "Return a JSON array of exactly " & fs.Count & " strings."

    Dim raw As String
    raw = ChatComplete(sys, user, provider, model)
    If Left$(raw, 6) = "Error:" Then FIELDS = raw: Exit Function

    Dim vals As Collection
    Set vals = ParseJsonStringArray(raw)

    Dim arr() As Variant
    ReDim arr(1 To 1, 1 To fs.Count)
    If Not vals Is Nothing And vals.Count = fs.Count Then
        For i = 1 To fs.Count
            arr(1, i) = Trim(CStr(vals(i)))
        Next i
    Else
        ' Fallback: extract each field individually.
        For i = 1 To fs.Count
            arr(1, i) = EXTRACT(text, fs(i), provider, model)
        Next i
    End If
    FIELDS = arr
    Exit Function
Fail:
    FIELDS = "Error: " & Err.Description
End Function

' =MAP(range, instruction, ...) -> applies the instruction to each cell (per cell).
Public Function MAP(rng As Variant, instruction As String, Optional provider As String = "", Optional model As String = "") As Variant
    On Error GoTo Fail
    Dim data As Variant
    data = ToArray2D(rng)

    Dim rows As Long, cols As Long
    rows = UBound(data, 1) - LBound(data, 1) + 1
    cols = UBound(data, 2) - LBound(data, 2) + 1

    Dim result() As Variant
    ReDim result(1 To rows, 1 To cols)

    Dim sys As String
    sys = "Apply the user's instruction to the single input value. Output only the result for that value, as plain text."

    Dim r As Long, c As Long, rr As Long, cc As Long, cell As String
    rr = 0
    For r = LBound(data, 1) To UBound(data, 1)
        rr = rr + 1: cc = 0
        For c = LBound(data, 2) To UBound(data, 2)
            cc = cc + 1
            cell = CStr(data(r, c))
            If Trim(cell) = "" Then
                result(rr, cc) = ""
            Else
                result(rr, cc) = TrimResult(ChatComplete(sys, "Instruction: " & instruction & vbLf & vbLf & "Input: " & cell, provider, model))
            End If
        Next c
    Next r
    MAP = result
    Exit Function
Fail:
    MAP = "Error: " & Err.Description
End Function

' =SIMILARITY(a, b, embeddingModel, [provider], [model]) -> cosine similarity.
Public Function SIMILARITY(a As String, b As String, embModel As String, Optional provider As String = "", Optional model As String = "") As Variant
    On Error GoTo Fail
    If embModel = "" Then SIMILARITY = "Error: pass an embedding model": Exit Function
    Dim va As Object, vb As Object
    Set va = EmbedVector(a, embModel, provider)
    Set vb = EmbedVector(b, embModel, provider)
    If va Is Nothing Or vb Is Nothing Then SIMILARITY = "Error: embedding failed": Exit Function
    SIMILARITY = Cosine(va, vb)
    Exit Function
Fail:
    SIMILARITY = "Error: " & Err.Description
End Function

' ---- helpers ----------------------------------------------------------------

Public Function MatchCategory(output As String, cats As Collection) As String
    Dim lower As String, i As Long
    lower = LCase(Trim(output))
    For i = 1 To cats.Count
        If LCase(cats(i)) = lower Then MatchCategory = cats(i): Exit Function
    Next i
    For i = 1 To cats.Count
        If InStr(lower, LCase(cats(i))) > 0 Then MatchCategory = cats(i): Exit Function
    Next i
    MatchCategory = Trim(output)
End Function

Public Function Cosine(a As Object, b As Object) As Double
    Dim n As Long, i As Long
    Dim dot As Double, na As Double, nb As Double, x As Double, y As Double
    n = a.Count
    If b.Count < n Then n = b.Count
    For i = 1 To n
        x = CDbl(a(i)): y = CDbl(b(i))
        dot = dot + x * y
        na = na + x * x
        nb = nb + y * y
    Next i
    If na = 0 Or nb = 0 Then Cosine = 0 Else Cosine = dot / (Sqr(na) * Sqr(nb))
End Function

' Number of dimensions of a Variant array (0 if not an array).
Private Function NumDims(v As Variant) As Long
    Dim n As Long, t As Long
    If Not IsArray(v) Then NumDims = 0: Exit Function
    On Error GoTo Done
    Do
        n = n + 1
        t = UBound(v, n)
    Loop
Done:
    NumDims = n - 1
End Function

' Flatten a scalar / 1D / 2D / range Variant into a Collection of trimmed,
' non-empty strings. Single cells containing "a, b, c" are comma-split.
Public Function FlattenToStrings(v As Variant) As Collection
    Dim c As New Collection
    Dim dims As Long
    dims = NumDims(v)
    If dims = 0 Then
        AddSplit c, CStr(v)
    ElseIf dims = 1 Then
        Dim i As Long
        For i = LBound(v) To UBound(v)
            AddSplit c, CStr(v(i))
        Next i
    Else
        Dim r As Long, col As Long
        For r = LBound(v, 1) To UBound(v, 1)
            For col = LBound(v, 2) To UBound(v, 2)
                AddSplit c, CStr(v(r, col))
            Next col
        Next r
    End If
    Set FlattenToStrings = c
End Function

Private Sub AddSplit(c As Collection, s As String)
    s = Trim(s)
    If s = "" Then Exit Sub
    If InStr(s, ",") > 0 Then
        Dim parts() As String, k As Long
        parts = Split(s, ",")
        For k = LBound(parts) To UBound(parts)
            If Trim(parts(k)) <> "" Then c.Add Trim(parts(k))
        Next k
    Else
        c.Add s
    End If
End Sub

' Flatten a Variant to text: rows by newline, columns by tab.
Public Function FlattenToText(v As Variant) As String
    Dim dims As Long
    dims = NumDims(v)
    If dims = 0 Then FlattenToText = CStr(v): Exit Function

    Dim s As String, i As Long, j As Long
    If dims = 2 Then
        For i = LBound(v, 1) To UBound(v, 1)
            Dim rowS As String
            rowS = ""
            For j = LBound(v, 2) To UBound(v, 2)
                rowS = rowS & IIf(j > LBound(v, 2), vbTab, "") & CStr(v(i, j))
            Next j
            s = s & IIf(i > LBound(v, 1), vbLf, "") & rowS
        Next i
    Else
        For i = LBound(v) To UBound(v)
            s = s & IIf(i > LBound(v), vbLf, "") & CStr(v(i))
        Next i
    End If
    FlattenToText = s
End Function

' Normalize a scalar / 1D / 2D Variant to a 2D array.
Private Function ToArray2D(v As Variant) As Variant
    Dim dims As Long
    dims = NumDims(v)
    If dims = 2 Then
        ToArray2D = v
    ElseIf dims = 1 Then
        Dim b() As Variant, i As Long, n As Long
        n = UBound(v) - LBound(v) + 1
        ReDim b(1 To 1, 1 To n)
        For i = LBound(v) To UBound(v)
            b(1, i - LBound(v) + 1) = v(i)
        Next i
        ToArray2D = b
    Else
        Dim a(1 To 1, 1 To 1) As Variant
        a(1, 1) = v
        ToArray2D = a
    End If
End Function

' Parse a JSON string array from model output (tolerates ``` fences). Returns a
' Collection of strings, or Nothing if it isn't a clean array.
Public Function ParseJsonStringArray(raw As String) As Collection
    On Error GoTo Fail
    Dim s As String
    s = Trim(raw)

    Dim p As Long, e As Long
    p = InStr(s, "```")
    If p > 0 Then
        s = Mid(s, p + 3)
        If LCase(Left(s, 4)) = "json" Then s = Mid(s, 5)
        e = InStr(s, "```")
        If e > 0 Then s = Left(s, e - 1)
        s = Trim(s)
    End If

    Dim st As Long, en As Long
    st = InStr(s, "[")
    en = InStrRev(s, "]")
    If st = 0 Or en <= st Then Set ParseJsonStringArray = Nothing: Exit Function

    Dim arr As Object
    Set arr = JsonConverter.ParseJson(Mid(s, st, en - st + 1))

    Dim c As New Collection, i As Long
    For i = 1 To arr.Count
        c.Add CStr(arr(i))
    Next i
    Set ParseJsonStringArray = c
    Exit Function
Fail:
    Set ParseJsonStringArray = Nothing
End Function

' Fallback list parsing: split lines, strip bullet/number prefixes.
Public Function SplitLinesToItems(raw As String) As Collection
    Dim c As New Collection
    Dim lines() As String, i As Long, ln As String
    lines = Split(Replace(raw, vbCr, ""), vbLf)
    For i = LBound(lines) To UBound(lines)
        ln = StripBullet(Trim(lines(i)))
        If ln <> "" Then c.Add ln
    Next i
    Set SplitLinesToItems = c
End Function

Private Function StripBullet(s As String) As String
    Dim t As String, ch As String, k As Long
    t = s
    Do While Len(t) > 0
        ch = Left(t, 1)
        If ch = "-" Or ch = "*" Or ch = Chr(8226) Or ch = Chr(149) Then
            t = Trim(Mid(t, 2))
        ElseIf ch Like "#" Then
            k = 1
            Do While k <= Len(t) And Mid(t, k, 1) Like "#"
                k = k + 1
            Loop
            If k <= Len(t) And (Mid(t, k, 1) = "." Or Mid(t, k, 1) = ")") Then
                t = Trim(Mid(t, k + 1))
            Else
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop
    StripBullet = t
End Function

Private Function TrimResult(s As String) As String
    If Left$(s, 6) = "Error:" Then TrimResult = s Else TrimResult = Trim(s)
End Function
