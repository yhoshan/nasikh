Attribute VB_Name = "NasikhTurath"
'============================================================
'  Nasikh (Turath edition) - Word macro to search turath.io
'  and insert text with reference  <book-name part/page>.
'
'  API:  https://api.turath.io/search?q=...&ver=3   (no key)
'  Needs an internet connection.
'
'  Run macro:  NasikhSearchTurath
'
'  turath search result item fields:
'     book_id, author_id, cat_id, snip, text,
'     meta : a JSON *string* containing
'            { book_name, author_name, vol, page, page_id }
'
'  Arabic UI texts use ChrW() so they are encoding-safe.
'============================================================
Option Explicit

Private Const API_URL As String = "https://api.turath.io/search"


'------------------------------------------------------------
' Entry point.
'------------------------------------------------------------
Public Sub NasikhSearchTurath()
    Dim query As String
    query = InputBox(T_PROMPT(), T_TITLE())
    If Trim(query) = "" Then Exit Sub

    Dim response As String
    On Error GoTo HttpError
    response = TurathSearch(query)
    On Error GoTo 0

    If Len(response) = 0 Then
        MsgBox T_NO_RESULTS_NET(), vbExclamation, T_NAME()
        Exit Sub
    End If

    InsertResults response, query
    Exit Sub

HttpError:
    MsgBox T_CONN_FAIL() & vbCrLf & Err.Description, vbCritical, T_NAME()
End Sub


'------------------------------------------------------------
' DIAGNOSTIC: writes raw first result slice into the document.
'------------------------------------------------------------
Public Sub NasikhDiagnoseTurath()
    Dim query As String
    query = InputBox(T_PROMPT(), T_TITLE())
    If Trim(query) = "" Then Exit Sub

    Dim response As String
    On Error GoTo HttpError
    response = TurathSearch(query)
    On Error GoTo 0

    Selection.TypeText "=== RAW (turath) ==="
    Selection.TypeParagraph
    Selection.TypeText Left$(response, 1800)
    Selection.TypeParagraph
    Exit Sub

HttpError:
    MsgBox T_CONN_FAIL() & vbCrLf & Err.Description, vbCritical, T_NAME()
End Sub


'------------------------------------------------------------
' GET request to turath search API.
'------------------------------------------------------------
Private Function TurathSearch(ByVal query As String) As String
    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    Dim url As String
    url = API_URL & "?q=" & UrlEncode(query) & "&ver=3"

    http.Open "GET", url, False
    http.setRequestHeader "Accept", "application/json"
    http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
    http.send

    If http.Status <> 200 Then
        Err.Raise vbObjectError + 1, , "HTTP " & http.Status
    End If

    TurathSearch = http.responseText
End Function


'------------------------------------------------------------
' Parse results and insert each (text + reference).
'------------------------------------------------------------
Private Sub InsertResults(ByVal json As String, ByVal query As String)
    Dim sel As Object
    Set sel = Selection

    sel.TypeText T_RESULTS_FOR() & query & T_QUOTE_CLOSE()
    sel.TypeParagraph
    sel.TypeParagraph

    ' Each result item begins with "snip" (or "text"); split on them.
    Dim items() As String
    items = SplitResults(json)

    Dim i As Long, n As Long
    n = 0
    For i = LBound(items) To UBound(items)
        Dim chunk As String
        chunk = items(i)

        ' body text: prefer snip (highlighted), else text
        Dim bodyText As String
        bodyText = CleanText(GetJsonValue(chunk, "snip"))
        If bodyText = "" Then bodyText = CleanText(GetJsonValue(chunk, "text"))

        ' meta is a JSON string (JSON inside JSON). GetJsonValue already
        ' unescaped one level (\" -> "), so now read fields directly.
        ' As a safety net we also handle a remaining escaped level.
        Dim meta As String
        meta = GetJsonValue(chunk, "meta")

        Dim bookName As String, pageNum As String, vol As String
        bookName = CleanText(GetMetaValue(meta, "book_name"))
        pageNum = GetMetaValue(meta, "page")
        vol = GetMetaValue(meta, "vol")

        If bookName <> "" Or bodyText <> "" Then
            n = n + 1
            sel.TypeText bodyText
            sel.TypeParagraph
            Dim ref As String
            ref = BuildReference(bookName, vol, pageNum)
            InsertReferenceRight sel, ref
            sel.TypeParagraph
            sel.TypeParagraph
        End If
    Next i

    If n = 0 Then
        sel.TypeText T_NO_MATCH()
        sel.TypeParagraph
    End If
End Sub


'------------------------------------------------------------
' Build reference:  <book part/page>  (part only if numeric)
'------------------------------------------------------------
Private Function BuildReference(ByVal name As String, ByVal vol As String, ByVal page As String) As String
    If name = "" Then name = T_UNKNOWN_BOOK()
    Dim loc As String

    If IsNumericClean(vol) And vol <> "0" Then
        If page <> "" Then
            loc = vol & "/" & page
        Else
            loc = vol
        End If
    ElseIf page <> "" Then
        loc = page
    End If

    If loc <> "" Then
        BuildReference = "<" & name & " " & loc & ">"
    Else
        BuildReference = "<" & name & ">"
    End If
End Function


Private Sub InsertReferenceRight(ByVal sel As Object, ByVal ref As String)
    sel.ParagraphFormat.Alignment = wdAlignParagraphRight
    sel.Font.Bold = True
    sel.TypeText ChrW(&H200F) & ref
    sel.Font.Bold = False
    sel.ParagraphFormat.Alignment = wdAlignParagraphRight
End Sub


'============================================================
'  Helpers (JSON reader + text cleaner)
'============================================================

Private Function SplitResults(ByVal json As String) As String()
    Dim marker As String
    marker = """snip"""
    If InStr(json, marker) = 0 Then marker = """text"""

    Dim parts() As String
    parts = Split(json, marker)

    Dim res() As String
    Dim cnt As Long
    cnt = UBound(parts) - LBound(parts)
    If cnt < 1 Then
        ReDim res(0 To 0)
        res(0) = json
        SplitResults = res
        Exit Function
    End If

    ReDim res(0 To cnt - 1)
    Dim k As Long
    For k = 1 To UBound(parts)
        res(k - 1) = marker & parts(k)
    Next k
    SplitResults = res
End Function


' Reads a simple "key":"value" or "key":value. Unescapes \" \\ \/ \uXXXX etc.
Private Function GetJsonValue(ByVal json As String, ByVal key As String) As String
    Dim pat As String
    pat = """" & key & """"
    Dim i As Long
    i = InStr(json, pat)
    If i = 0 Then Exit Function

    i = InStr(i + Len(pat), json, ":")
    If i = 0 Then Exit Function
    i = i + 1

    Do While i <= Len(json) And Mid$(json, i, 1) = " "
        i = i + 1
    Loop
    If i > Len(json) Then Exit Function

    Dim ch As String
    ch = Mid$(json, i, 1)
    If ch = """" Then
        Dim j As Long
        j = i + 1
        Dim out As String
        Do While j <= Len(json)
            Dim c As String
            c = Mid$(json, j, 1)
            If c = "\" Then
                Dim nx As String
                nx = Mid$(json, j + 1, 1)
                Select Case nx
                    Case "n": out = out & vbCrLf
                    Case "r": out = out & ""
                    Case "t": out = out & " "
                    Case """": out = out & """"
                    Case "\": out = out & "\"
                    Case "/": out = out & "/"
                    Case "u"
                        Dim hx As String
                        hx = Mid$(json, j + 2, 4)
                        out = out & ChrW(CLng("&H" & hx))
                        j = j + 4
                    Case Else: out = out & nx
                End Select
                j = j + 2
            ElseIf c = """" Then
                Exit Do
            Else
                out = out & c
                j = j + 1
            End If
        Loop
        GetJsonValue = out
    Else
        Dim m As Long
        m = i
        Dim num As String
        Do While m <= Len(json)
            Dim d As String
            d = Mid$(json, m, 1)
            If d = "," Or d = "}" Or d = "]" Then Exit Do
            num = num & d
            m = m + 1
        Loop
        GetJsonValue = Trim$(num)
    End If
End Function


' يقرأ قيمة من نص meta سواء كانت علامات الاقتباس عادية (") أو مهرّبة (\").
' يتعامل مع:   "key":"value"   و   \"key\":\"value\"   و   \"key\":value
Private Function GetMetaValue(ByVal meta As String, ByVal key As String) As String
    ' أولًا: جرّب القراءة العادية (علامات اقتباس غير مهرّبة)
    Dim v As String
    v = GetJsonValue(meta, key)
    If v <> "" Then
        GetMetaValue = v
        Exit Function
    End If

    ' ثانيًا: ابحث عن المفتاح المهرّب  \"key\"
    Dim pat As String
    pat = "\""" & key & "\"""      ' = \"key\"
    Dim i As Long
    i = InStr(meta, pat)
    If i = 0 Then Exit Function

    i = InStr(i + Len(pat), meta, ":")
    If i = 0 Then Exit Function
    i = i + 1
    Do While i <= Len(meta) And Mid$(meta, i, 1) = " "
        i = i + 1
    Loop
    If i > Len(meta) Then Exit Function

    ' القيمة قد تبدأ بـ  \"  (نص)  أو برقم
    If Mid$(meta, i, 2) = "\""" Then
        ' نص مهرّب: اقرأ حتى  \"  التالية
        Dim j As Long
        j = i + 2
        Dim out As String
        Do While j <= Len(meta)
            If Mid$(meta, j, 2) = "\""" Then Exit Do
            out = out & Mid$(meta, j, 1)
            j = j + 1
        Loop
        GetMetaValue = out
    Else
        ' رقم أو قيمة بسيطة: اقرأ حتى , أو } أو \
        Dim m As Long
        m = i
        Dim num As String
        Do While m <= Len(meta)
            Dim d As String
            d = Mid$(meta, m, 1)
            If d = "," Or d = "}" Or d = "\" Or d = "]" Then Exit Do
            num = num & d
            m = m + 1
        Loop
        GetMetaValue = Trim$(num)
    End If
End Function


Private Function CleanText(ByVal s As String) As String
    If s = "" Then Exit Function
    Dim out As String, inTag As Boolean, k As Long, c As String
    For k = 1 To Len(s)
        c = Mid$(s, k, 1)
        If c = "<" Then
            inTag = True
        ElseIf c = ">" Then
            inTag = False
        ElseIf Not inTag Then
            out = out & c
        End If
    Next k
    out = Replace(out, "&amp;", "&")
    out = Replace(out, "&lt;", "<")
    out = Replace(out, "&gt;", ">")
    out = Replace(out, "&quot;", """")
    Do While InStr(out, "  ") > 0
        out = Replace(out, "  ", " ")
    Loop
    CleanText = Trim$(out)
End Function


Private Function IsNumericClean(ByVal s As String) As Boolean
    s = Trim$(s)
    If s = "" Then Exit Function
    Dim k As Long, c As String
    For k = 1 To Len(s)
        c = Mid$(s, k, 1)
        If Not (c >= "0" And c <= "9") And Not (AscW(c) >= &H660 And AscW(c) <= &H669) Then
            Exit Function
        End If
    Next k
    IsNumericClean = True
End Function


' URL-encode (UTF-8 percent encoding) for the query string.
Private Function UrlEncode(ByVal s As String) As String
    Dim i As Long, c As Long, out As String
    Dim bytes() As Byte
    ' convert to UTF-8 bytes
    bytes = Utf8Bytes(s)
    For i = 0 To UBound(bytes)
        c = bytes(i)
        If (c >= 48 And c <= 57) Or (c >= 65 And c <= 90) Or (c >= 97 And c <= 122) _
           Or c = 45 Or c = 46 Or c = 95 Or c = 126 Then
            out = out & Chr$(c)
        Else
            out = out & "%" & Right$("0" & Hex$(c), 2)
        End If
    Next i
    UrlEncode = out
End Function


' Returns UTF-8 bytes of a (possibly Arabic) string.
Private Function Utf8Bytes(ByVal s As String) As Byte()
    Dim i As Long, cp As Long, b() As Byte, n As Long
    ReDim b(0 To Len(s) * 4)   ' upper bound
    n = 0
    For i = 1 To Len(s)
        cp = AscW(Mid$(s, i, 1))
        If cp < 0 Then cp = cp + 65536
        If cp < &H80 Then
            b(n) = cp: n = n + 1
        ElseIf cp < &H800 Then
            b(n) = &HC0 Or (cp \ &H40): n = n + 1
            b(n) = &H80 Or (cp And &H3F): n = n + 1
        Else
            b(n) = &HE0 Or (cp \ &H1000): n = n + 1
            b(n) = &H80 Or ((cp \ &H40) And &H3F): n = n + 1
            b(n) = &H80 Or (cp And &H3F): n = n + 1
        End If
    Next i
    If n = 0 Then
        ReDim b(0 To 0)
        Utf8Bytes = b
        Exit Function
    End If
    ReDim Preserve b(0 To n - 1)
    Utf8Bytes = b
End Function


'============================================================
'  Arabic UI texts via ChrW (encoding-safe)
'============================================================

' اكتب عبارة البحث في مكتبة التراث:
Private Function T_PROMPT() As String
    T_PROMPT = ChrW(1575) & ChrW(1603) & ChrW(1578) & ChrW(1576) & ChrW(32) & ChrW(1593) & ChrW(1576) & ChrW(1575) & ChrW(1585) & ChrW(1577) & ChrW(32) & ChrW(1575) & ChrW(1604) & ChrW(1576) & ChrW(1581) & ChrW(1579) & ChrW(32) & ChrW(1601) & ChrW(1610) & ChrW(32) & ChrW(1605) & ChrW(1603) & ChrW(1578) & ChrW(1576) & ChrW(1577) & ChrW(32) & ChrW(1575) & ChrW(1604) & ChrW(1578) & ChrW(1585) & ChrW(1575) & ChrW(1579) & ChrW(58)
End Function

' ناسخ — بحث
Private Function T_TITLE() As String
    T_TITLE = ChrW(1606) & ChrW(1575) & ChrW(1587) & ChrW(1582) & ChrW(32) & ChrW(8212) & ChrW(32) & ChrW(1576) & ChrW(1581) & ChrW(1579)
End Function

' ناسخ
Private Function T_NAME() As String
    T_NAME = ChrW(1606) & ChrW(1575) & ChrW(1587) & ChrW(1582)
End Function

' لم تصل نتائج. تأكد من الاتصال بالإنترنت.
Private Function T_NO_RESULTS_NET() As String
    T_NO_RESULTS_NET = ChrW(1604) & ChrW(1605) & ChrW(32) & ChrW(1578) & ChrW(1589) & ChrW(1604) & ChrW(32) & ChrW(1606) & ChrW(1578) & ChrW(1575) & ChrW(1574) & ChrW(1580) & ChrW(46) & ChrW(32) & ChrW(1578) & ChrW(1571) & ChrW(1603) & ChrW(1583) & ChrW(32) & ChrW(1605) & ChrW(1606) & ChrW(32) & ChrW(1575) & ChrW(1604) & ChrW(1575) & ChrW(1578) & ChrW(1589) & ChrW(1575) & ChrW(1604) & ChrW(32) & ChrW(1576) & ChrW(1575) & ChrW(1604) & ChrW(1573) & ChrW(1606) & ChrW(1578) & ChrW(1585) & ChrW(1606) & ChrW(1578) & ChrW(46)
End Function

' تعذّر الاتصال بالتراث:
Private Function T_CONN_FAIL() As String
    T_CONN_FAIL = ChrW(1578) & ChrW(1593) & ChrW(1584) & ChrW(1617) & ChrW(1585) & ChrW(32) & ChrW(1575) & ChrW(1604) & ChrW(1575) & ChrW(1578) & ChrW(1589) & ChrW(1575) & ChrW(1604) & ChrW(32) & ChrW(1576) & ChrW(1575) & ChrW(1604) & ChrW(1578) & ChrW(1585) & ChrW(1575) & ChrW(1579) & ChrW(58)
End Function

' نتائج البحث عن: «
Private Function T_RESULTS_FOR() As String
    T_RESULTS_FOR = ChrW(1606) & ChrW(1578) & ChrW(1575) & ChrW(1574) & ChrW(1580) & ChrW(32) & ChrW(1575) & ChrW(1604) & ChrW(1576) & ChrW(1581) & ChrW(1579) & ChrW(32) & ChrW(1593) & ChrW(1606) & ChrW(58) & ChrW(32) & ChrW(171)
End Function

' »
Private Function T_QUOTE_CLOSE() As String
    T_QUOTE_CLOSE = ChrW(187)
End Function

' لا توجد نتائج مطابقة.
Private Function T_NO_MATCH() As String
    T_NO_MATCH = ChrW(1604) & ChrW(1575) & ChrW(32) & ChrW(1578) & ChrW(1608) & ChrW(1580) & ChrW(1583) & ChrW(32) & ChrW(1606) & ChrW(1578) & ChrW(1575) & ChrW(1574) & ChrW(1580) & ChrW(32) & ChrW(1605) & ChrW(1591) & ChrW(1575) & ChrW(1576) & ChrW(1602) & ChrW(1577) & ChrW(46)
End Function

' كتاب غير معروف
Private Function T_UNKNOWN_BOOK() As String
    T_UNKNOWN_BOOK = ChrW(1603) & ChrW(1578) & ChrW(1575) & ChrW(1576) & ChrW(32) & ChrW(1594) & ChrW(1610) & ChrW(1585) & ChrW(32) & ChrW(1605) & ChrW(1593) & ChrW(1585) & ChrW(1608) & ChrW(1601)
End Function
