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

' =TAG(text, categories, ...) -> comma-separated matching labels (multi-label).
Public Function TAG(text As String, categories As Variant, Optional provider As String = "", Optional model As String = "") As String
    On Error GoTo Fail
    Dim cats As Collection
    Set cats = FlattenToStrings(categories)
    If cats.Count = 0 Then TAG = "Error: no categories provided": Exit Function

    Dim labels As String, i As Long
    For i = 1 To cats.Count
        labels = labels & IIf(i > 1, ", ", "") & cats(i)
    Next i

    Dim sys As String
    sys = "Apply labels to the text. Choose ALL that apply from: " & labels & ". " & _
          "Return only the matching labels as a comma-separated list, in the given order. If none apply, return nothing."
    Dim out As String
    out = ChatComplete(sys, "Text:" & vbLf & text, provider, model)
    If Left$(out, 6) = "Error:" Then TAG = out: Exit Function

    ' Keep only recognized labels (order-preserving) so the result is always clean.
    Dim lc As String, res As String
    lc = LCase$(out)
    For i = 1 To cats.Count
        If InStr(lc, LCase$(cats(i))) > 0 Then res = res & IIf(res = "", "", ", ") & cats(i)
    Next i
    TAG = res
    Exit Function
Fail:
    TAG = "Error: " & Err.Description
End Function

' =EDIT(text, [instruction], ...) -> revised text (default: fix spelling & grammar).
Public Function EDIT(text As String, Optional instruction As String = "", Optional provider As String = "", Optional model As String = "") As String
    Dim what As String
    what = Trim$(instruction)
    If what = "" Then what = "Fix spelling and grammar"
    Dim sys As String
    sys = "You are an editor. Apply the requested edit and output ONLY the revised text -- no notes or quotes."
    EDIT = TrimResult(ChatComplete(sys, "Edit instruction: " & what & vbLf & vbLf & "Text:" & vbLf & text, provider, model))
End Function

' =FORMULA(description, ...) -> an Excel formula string.
Public Function FORMULA(description As String, Optional provider As String = "", Optional model As String = "") As String
    Dim sys As String
    sys = "You write Microsoft Excel formulas. Output ONLY a single Excel formula that " & _
          "starts with '=' -- no explanation, no code fences, no surrounding text. Prefer " & _
          "standard, widely-supported functions."
    Dim out As String
    out = ChatComplete(sys, "Write an Excel formula that: " & description, provider, model)
    If Left$(out, 6) = "Error:" Then FORMULA = out: Exit Function
    FORMULA = CleanFormula(out)
End Function

' =EXPLAIN(formula, ...) -> plain-English explanation. Tip: =EXPLAIN(FORMULATEXT(A1)).
Public Function EXPLAIN(formula As String, Optional provider As String = "", Optional model As String = "") As String
    Dim sys As String
    sys = "Explain what the given Excel formula does, in plain English and concisely " & _
          "(2-3 sentences max). Output only the explanation."
    EXPLAIN = TrimResult(ChatComplete(sys, "Explain this Excel formula:" & vbLf & formula, provider, model))
End Function

' =LLMTABLE(prompt, ...) -> spills a 2D grid (first row = headers). Named LLMTABLE
' because a bare TABLE collides with Excel's legacy Data Table function.
Public Function LLMTABLE(promptText As String, Optional provider As String = "", Optional model As String = "") As Variant
    On Error GoTo Fail
    Dim sys As String
    sys = "Generate tabular data as a JSON array of arrays: each inner array is one row " & _
          "of string cells, and the first row is the column headers. Keep every row the " & _
          "same length. Return ONLY the JSON -- no commentary or code fences."
    Dim raw As String
    raw = ChatComplete(sys, promptText, provider, model)
    If Left$(raw, 6) = "Error:" Then LLMTABLE = raw: Exit Function
    Dim grid As Variant
    grid = ParseJsonGrid(raw)
    If IsEmpty(grid) Then LLMTABLE = "Error: could not parse a table from the response": Exit Function
    LLMTABLE = grid
    Exit Function
Fail:
    LLMTABLE = "Error: " & Err.Description
End Function

' =FILL(examples, inputs, ...) -> infer a pattern from (input,output) example pairs
' and apply it to new inputs. examples is a two-column range; inputs is a range.
Public Function FILL(examples As Variant, inputs As Variant, Optional provider As String = "", Optional model As String = "") As Variant
    On Error GoTo Fail

    ' Collect example (input, output) pairs into a numbered block.
    Dim exBlock As String, nEx As Long
    If IsArray(examples) Then
        Dim er As Long, c0 As Long
        c0 = LBound(examples, 2)
        For er = LBound(examples, 1) To UBound(examples, 1)
            Dim ei As String, eo As String
            ei = Trim$(CStr(examples(er, c0)))
            eo = Trim$(CStr(examples(er, c0 + 1)))
            If ei <> "" And eo <> "" Then
                nEx = nEx + 1
                exBlock = exBlock & nEx & ". IN: " & ei & "  =>  OUT: " & eo & vbLf
            End If
        Next er
    End If
    If nEx = 0 Then FILL = "Error: need at least one example (input, output) pair": Exit Function

    ' Flatten inputs (row-major) and remember the shape for the output.
    Dim flat As Collection: Set flat = New Collection
    Dim isArr As Boolean: isArr = IsArray(inputs)
    Dim rr As Long, cc As Long
    If isArr Then
        For rr = LBound(inputs, 1) To UBound(inputs, 1)
            For cc = LBound(inputs, 2) To UBound(inputs, 2)
                flat.Add CStr(inputs(rr, cc))
            Next cc
        Next rr
    Else
        flat.Add CStr(inputs)
    End If
    If flat.Count = 0 Then FILL = "": Exit Function

    ' One batched call, JSON array of outputs.
    Dim inBlock As String, i As Long
    For i = 1 To flat.Count
        inBlock = inBlock & i & ". " & flat(i) & vbLf
    Next i
    Dim sys As String
    sys = "Infer the transformation from the examples and apply it to each new input. " & _
          "Return ONLY a JSON array of output strings, exactly one per input, in order. No commentary."
    Dim user As String
    user = "Examples:" & vbLf & exBlock & vbLf & _
           "Apply the same transformation to these inputs and return a JSON array of exactly " & _
           flat.Count & " strings:" & vbLf & inBlock
    Dim raw As String
    raw = ChatComplete(sys, user, provider, model)

    Dim results As Collection
    Set results = Nothing
    If Left$(raw, 6) <> "Error:" Then Set results = ParseJsonStringArray(raw)
    If results Is Nothing Then
        Set results = FillPerInput(exBlock, flat, provider, model)
    ElseIf results.Count <> flat.Count Then
        Set results = FillPerInput(exBlock, flat, provider, model)
    End If

    ' Reshape results back to the inputs' shape.
    If isArr Then
        Dim outArr() As String
        ReDim outArr(LBound(inputs, 1) To UBound(inputs, 1), LBound(inputs, 2) To UBound(inputs, 2))
        Dim k As Long: k = 1
        For rr = LBound(inputs, 1) To UBound(inputs, 1)
            For cc = LBound(inputs, 2) To UBound(inputs, 2)
                outArr(rr, cc) = CStr(results(k))
                k = k + 1
            Next cc
        Next rr
        FILL = outArr
    Else
        FILL = CStr(results(1))
    End If
    Exit Function
Fail:
    FILL = "Error: " & Err.Description
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

' =RECALL(query, candidates, [k], [embModel], [provider], [model]) -> top-k rows
' from a range ranked by embedding similarity to the query, as [text, score].
Public Function RECALL(query As String, candidates As Variant, Optional k As Long = 5, _
                       Optional embModel As String = "", Optional provider As String = "", _
                       Optional model As String = "") As Variant
    On Error GoTo Fail
    If embModel = "" Then RECALL = "Error: pass an embedding model": Exit Function

    ' Flatten candidates to non-empty strings.
    Dim texts As Collection: Set texts = New Collection
    If IsArray(candidates) Then
        Dim r As Long, c As Long, t As String
        For r = LBound(candidates, 1) To UBound(candidates, 1)
            For c = LBound(candidates, 2) To UBound(candidates, 2)
                t = Trim$(CStr(candidates(r, c)))
                If t <> "" Then texts.Add t
            Next c
        Next r
    Else
        Dim s As String: s = Trim$(CStr(candidates))
        If s <> "" Then texts.Add s
    End If
    If texts.Count = 0 Then RECALL = "Error: no candidates": Exit Function

    Dim qv As Object
    Set qv = EmbedVector(query, embModel, provider)
    If qv Is Nothing Then RECALL = "Error: embedding failed": Exit Function

    Dim n As Long, i As Long
    n = texts.Count
    Dim scores() As Double, keep() As String
    ReDim scores(1 To n)
    ReDim keep(1 To n)
    For i = 1 To n
        keep(i) = texts(i)
        Dim cv As Object
        Set cv = EmbedVector(texts(i), embModel, provider)
        If cv Is Nothing Then scores(i) = -1 Else scores(i) = Cosine(qv, cv)
    Next i

    ' Selection-pick the top-k by descending score.
    Dim topN As Long
    topN = k
    If topN <= 0 Or topN > n Then topN = n
    Dim outArr() As Variant
    ReDim outArr(1 To topN, 1 To 2)
    Dim picked As Long, j As Long, best As Long
    Dim bestScore As Double
    For picked = 1 To topN
        best = 0: bestScore = -2
        For j = 1 To n
            If scores(j) > bestScore Then bestScore = scores(j): best = j
        Next j
        If best = 0 Then Exit For
        outArr(picked, 1) = keep(best)
        outArr(picked, 2) = Application.Round(bestScore, 3)
        scores(best) = -2   ' mark as consumed
    Next picked
    RECALL = outArr
    Exit Function
Fail:
    RECALL = "Error: " & Err.Description
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

    Dim sliced As String
    sliced = Mid(s, st, en - st + 1)

    ' Small/local models often emit a trailing comma before ]; VBA-JSON rejects it.
    ' Try strict first, then a portable trailing-comma repair (no VBScript.RegExp,
    ' which is absent on Mac Excel). On failure the callers take their own reliable
    ' fallback, so a failed repair is never worse than the un-repaired behavior.
    Dim arr As Object
    Set arr = TryParseJsonArray(sliced)
    If arr Is Nothing Then Set arr = TryParseJsonArray(StripTrailingComma(sliced))
    If arr Is Nothing Then Set ParseJsonStringArray = Nothing: Exit Function

    Dim c As New Collection, i As Long
    For i = 1 To arr.Count
        c.Add CStr(arr(i))
    Next i
    Set ParseJsonStringArray = c
    Exit Function
Fail:
    Set ParseJsonStringArray = Nothing
End Function

' Pull a clean "=..." formula out of a model reply (strip fences/backticks/prose).
Private Function CleanFormula(ByVal raw As String) As String
    Dim s As String
    s = Trim$(raw)
    s = Replace(s, "```excel", "")
    s = Replace(s, "```json", "")
    s = Replace(s, "```vba", "")
    s = Replace(s, "```", "")
    s = Replace(s, "`", "")
    s = Trim$(s)
    Dim eq As Long
    eq = InStr(s, "=")
    If eq > 0 Then s = Mid$(s, eq)   ' drop any leading prose before '='
    Dim lf As Long
    lf = InStr(s, vbLf)
    If lf > 0 Then s = Left$(s, lf - 1)   ' formula is a single line
    lf = InStr(s, vbCr)
    If lf > 0 Then s = Left$(s, lf - 1)
    s = Trim$(s)
    If s <> "" And Left$(s, 1) <> "=" Then s = "=" & s
    CleanFormula = s
End Function

' Parse a JSON array-of-arrays into a 1-based 2D String array; Empty on failure.
Private Function ParseJsonGrid(ByVal raw As String) As Variant
    On Error GoTo Fail
    Dim s As String
    s = Trim$(raw)
    Dim p As Long
    p = InStr(s, "```")
    If p > 0 Then
        s = Mid$(s, p + 3)
        If LCase$(Left$(s, 4)) = "json" Then s = Mid$(s, 5)
        Dim ef As Long
        ef = InStr(s, "```")
        If ef > 0 Then s = Left$(s, ef - 1)
        s = Trim$(s)
    End If

    Dim st As Long, en As Long
    st = InStr(s, "[")
    en = InStrRev(s, "]")
    If st = 0 Or en <= st Then ParseJsonGrid = Empty: Exit Function

    Dim sliced As String
    sliced = Mid$(s, st, en - st + 1)
    Dim outer As Object
    Set outer = TryParseJsonArray(sliced)
    If outer Is Nothing Then Set outer = TryParseJsonArray(StripTrailingComma(sliced))
    If outer Is Nothing Then ParseJsonGrid = Empty: Exit Function
    If outer.Count = 0 Then ParseJsonGrid = Empty: Exit Function

    Dim nRows As Long, nCols As Long, r As Long, c As Long
    nRows = outer.Count

    ' VBA-JSON returns arrays as Collections; a row is itself a Collection.
    If IsObject(outer(1)) Then
        nCols = outer(1).Count
        Dim grid() As String
        ReDim grid(1 To nRows, 1 To nCols)
        For r = 1 To nRows
            Dim row As Object
            Set row = outer(r)
            For c = 1 To nCols
                If c <= row.Count Then grid(r, c) = CStr(row(c)) Else grid(r, c) = ""
            Next c
        Next r
        ParseJsonGrid = grid
    Else
        ' Flat array -> single column.
        Dim col() As String
        ReDim col(1 To nRows, 1 To 1)
        For r = 1 To nRows
            col(r, 1) = CStr(outer(r))
        Next r
        ParseJsonGrid = col
    End If
    Exit Function
Fail:
    ParseJsonGrid = Empty
End Function

' Fallback for FILL: apply the transformation one input at a time.
Private Function FillPerInput(ByVal exBlock As String, ByVal flat As Collection, _
                              ByVal provider As String, ByVal model As String) As Collection
    Dim c As New Collection, i As Long, one As String
    Dim sys As String
    sys = "Infer the transformation from the examples and output ONLY the result for the " & _
          "new input -- plain text, no labels or quotes."
    For i = 1 To flat.Count
        one = ChatComplete(sys, exBlock & "IN: " & flat(i) & "  =>  OUT:", provider, model)
        If Left$(one, 6) = "Error:" Then c.Add one Else c.Add Trim$(one)
    Next i
    Set FillPerInput = c
End Function

' Parse a JSON array, returning Nothing instead of raising on malformed input.
Private Function TryParseJsonArray(ByVal s As String) As Object
    On Error GoTo Fail
    Set TryParseJsonArray = JsonConverter.ParseJson(s)
    Exit Function
Fail:
    Set TryParseJsonArray = Nothing
End Function

' Remove a single trailing comma between the last element and the closing "]"
' (e.g. ["a","b",] -> ["a","b"]). Only touches the comma right before "]", so it
' cannot corrupt string contents. Pure VBA -- works on Mac and Windows.
Private Function StripTrailingComma(ByVal s As String) As String
    Dim rb As Long
    rb = InStrRev(s, "]")
    If rb = 0 Then StripTrailingComma = s: Exit Function

    Dim i As Long, ch As String
    For i = rb - 1 To 1 Step -1
        ch = Mid(s, i, 1)
        If ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Then
            ' skip whitespace between the comma and "]"
        ElseIf ch = "," Then
            s = Left(s, i - 1) & Mid(s, i + 1)
            Exit For
        Else
            Exit For   ' last real char isn't a comma; nothing to strip
        End If
    Next i
    StripTrailingComma = s
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
