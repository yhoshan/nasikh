Attribute VB_Name = "Nasikh"
'============================================================
'  Nasikh - Word macro to search Aljam3 library and insert
'  the text with reference  <book-name part/page>.
'
'  Connects to Aljam3 public MCP server over HTTPS (no key).
'  Needs an internet connection.
'
'  Run macro:  NasikhSearch   (or from the ribbon button).
'
'  NOTE: All Arabic UI texts are built via ChrW() helper
'  functions so they are NOT affected by file encoding.
'============================================================
Option Explicit

Private Const MCP_URL As String = "https://aljam3.com/mcp"


'------------------------------------------------------------
' DIAGNOSTIC: writes the RAW first result into the document so
' we can see the exact field names (book title / vol / page).
' Run this once (Alt+F8 > NasikhDiagnose) and send the output.
'------------------------------------------------------------
Public Sub NasikhDiagnose()
    Dim query As String
    query = InputBox(T_PROMPT(), T_TITLE())
    If Trim(query) = "" Then Exit Sub

    Dim response As String
    On Error GoTo HttpError
    response = McpSearch(query, 1)
    On Error GoTo 0

    Dim p As Long
    p = InStr(response, "data:")
    If p > 0 Then response = Mid$(response, p + 5)

    ' take a reasonable slice around the first result
    Dim startAt As Long
    startAt = InStr(response, """results""")
    If startAt = 0 Then startAt = 1
    Dim slice As String
    slice = Mid$(response, startAt, 1500)

    Selection.TypeText "=== RAW (first result) ==="
    Selection.TypeParagraph
    Selection.TypeText slice
    Selection.TypeParagraph
    Exit Sub

HttpError:
    MsgBox T_CONN_FAIL() & vbCrLf & Err.Description, vbCritical, T_NAME()
End Sub


'------------------------------------------------------------
' Entry point: ask for query, search, insert results.
'------------------------------------------------------------
Public Sub NasikhSearch()
    Dim query As String
    query = InputBox(T_PROMPT(), T_TITLE())
    If Trim(query) = "" Then Exit Sub

    Dim response As String
    On Error GoTo HttpError
    response = McpSearch(query, 10)
    On Error GoTo 0

    If Len(response) = 0 Then
        MsgBox T_NO_RESULTS_NET(), vbExclamation, T_NAME()
        Exit Sub
    End If

    InsertResultsFromJson response, query
    Exit Sub

HttpError:
    MsgBox T_CONN_FAIL() & vbCrLf & Err.Description, vbCritical, T_NAME()
End Sub


'------------------------------------------------------------
' Send MCP (JSON-RPC) request to the search tool.
'------------------------------------------------------------
Private Function McpSearch(ByVal query As String, ByVal perPage As Long) As String
    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    Dim body As String
    body = "{""jsonrpc"":""2.0"",""id"":1,""method"":""tools/call""," & _
           """params"":{""name"":""search"",""arguments"":{" & _
           """q"":""" & JsonEscape(query) & """,""per_page"":" & perPage & "}}}"

    http.Open "POST", MCP_URL, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Accept", "application/json, text/event-stream"
    http.send body

    If http.Status <> 200 Then
        Err.Raise vbObjectError + 1, , "HTTP " & http.Status
    End If

    McpSearch = http.responseText
End Function


'------------------------------------------------------------
' Parse the response and insert each result (text + reference).
'------------------------------------------------------------
Private Sub InsertResultsFromJson(ByVal json As String, ByVal query As String)
    ' If SSE stream, take what follows the first "data:"
    Dim p As Long
    p = InStr(json, "data:")
    If p > 0 Then json = Mid$(json, p + 5)

    Dim sel As Object
    Set sel = Selection

    ' Title
    sel.TypeText T_RESULTS_FOR() & query & T_QUOTE_CLOSE()
    sel.TypeParagraph
    sel.TypeParagraph

    Dim items() As String
    items = SplitResults(json)

    Dim i As Long, n As Long
    n = 0
    For i = LBound(items) To UBound(items)
        Dim chunk As String
        chunk = items(i)

        Dim excerpt As String, bookName As String, pageNum As String, vol As String
        excerpt = CleanText(GetJsonValue(chunk, "excerpt"))
        If excerpt = "" Then excerpt = CleanText(GetJsonValue(chunk, "snip"))
        If excerpt = "" Then excerpt = CleanText(GetJsonValue(chunk, "content"))

        bookName = CleanText(GetBookName(chunk))
        pageNum = GetJsonValue(chunk, "page_number")
        If pageNum = "" Then pageNum = GetJsonValue(chunk, "number")
        If pageNum = "" Then pageNum = GetJsonValue(chunk, "page")
        vol = GetJsonValue(chunk, "vol")
        If vol = "" Then vol = GetJsonValue(chunk, "volume")
        If vol = "" Then vol = GetJsonValue(chunk, "part")

        If bookName <> "" Or excerpt <> "" Then
            n = n + 1
            sel.TypeText excerpt
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
' Build reference:  <book part/page>
' Ignores part unless it is a clean number; else page only.
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


'------------------------------------------------------------
' Insert reference right-aligned (with RLM mark).
'------------------------------------------------------------
Private Sub InsertReferenceRight(ByVal sel As Object, ByVal ref As String)
    sel.ParagraphFormat.Alignment = wdAlignParagraphRight
    sel.Font.Bold = True
    sel.TypeText ChrW(&H200F) & ref
    sel.Font.Bold = False
    sel.ParagraphFormat.Alignment = wdAlignParagraphRight
End Sub


'============================================================
'  Helpers (simple JSON reader + text cleaner)
'============================================================

Private Function SplitResults(ByVal json As String) As String()
    Dim marker As String
    marker = """excerpt"""
    If InStr(json, marker) = 0 Then marker = """snip"""
    If InStr(json, marker) = 0 Then marker = """content"""

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


' يستخرج عنوان الكتاب فقط (يتجنّب الخلط مع اسم المؤلف).
' الأولوية: book_name > book.name/title (من كائن book) > title على المستوى الأعلى.
Private Function GetBookName(ByVal json As String) As String
    ' 1) حقل مباشر book_name
    GetBookName = GetJsonValue(json, "book_name")
    If GetBookName <> "" Then Exit Function

    ' 2) داخل كائن "book": { ... "name"/"title" ... } مع قصّ الكائن عند أول }
    Dim i As Long
    i = InStr(json, """book""")
    If i > 0 Then
        Dim rest As String
        rest = Mid$(json, i)
        Dim obj As String
        obj = ExtractObjectAfterKey(rest, "book")
        If obj <> "" Then
            GetBookName = GetJsonValue(obj, "name")
            If GetBookName <> "" Then Exit Function
            GetBookName = GetJsonValue(obj, "title")
            If GetBookName <> "" Then Exit Function
        End If
    End If

    ' 3) حقل title على المستوى الأعلى (قبل ظهور author لتفادي اسم المؤلف)
    Dim posTitle As Long, posAuthor As Long
    posTitle = InStr(json, """title""")
    posAuthor = InStr(json, """author""")
    If posTitle > 0 And (posAuthor = 0 Or posTitle < posAuthor) Then
        GetBookName = GetJsonValue(json, "title")
    End If
End Function


' يستخرج محتوى كائن JSON يلي مفتاحًا معيّنًا: "key":{ ... }
Private Function ExtractObjectAfterKey(ByVal json As String, ByVal key As String) As String
    Dim pat As String
    pat = """" & key & """"
    Dim i As Long
    i = InStr(json, pat)
    If i = 0 Then Exit Function
    i = InStr(i + Len(pat), json, "{")
    If i = 0 Then Exit Function

    Dim depth As Long, j As Long
    depth = 0
    For j = i To Len(json)
        Dim c As String
        c = Mid$(json, j, 1)
        If c = "{" Then
            depth = depth + 1
        ElseIf c = "}" Then
            depth = depth - 1
            If depth = 0 Then
                ExtractObjectAfterKey = Mid$(json, i, j - i + 1)
                Exit Function
            End If
        End If
    Next j
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


Private Function JsonEscape(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    JsonEscape = s
End Function


'============================================================
'  Arabic UI texts built via ChrW (encoding-safe)
'============================================================

' اكتب عبارة البحث في مكتبة الجامع:
Private Function T_PROMPT() As String
    T_PROMPT = ChrW(1575) & ChrW(1603) & ChrW(1578) & ChrW(1576) & ChrW(32) & ChrW(1593) & ChrW(1576) & ChrW(1575) & ChrW(1585) & ChrW(1577) & ChrW(32) & ChrW(1575) & ChrW(1604) & ChrW(1576) & ChrW(1581) & ChrW(1579) & ChrW(32) & ChrW(1601) & ChrW(1610) & ChrW(32) & ChrW(1605) & ChrW(1603) & ChrW(1578) & ChrW(1576) & ChrW(1577) & ChrW(32) & ChrW(1575) & ChrW(1604) & ChrW(1580) & ChrW(1575) & ChrW(1605) & ChrW(1593) & ChrW(58)
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

' تعذّر الاتصال بالجامع:
Private Function T_CONN_FAIL() As String
    T_CONN_FAIL = ChrW(1578) & ChrW(1593) & ChrW(1584) & ChrW(1617) & ChrW(1585) & ChrW(32) & ChrW(1575) & ChrW(1604) & ChrW(1575) & ChrW(1578) & ChrW(1589) & ChrW(1575) & ChrW(1604) & ChrW(32) & ChrW(1576) & ChrW(1575) & ChrW(1604) & ChrW(1580) & ChrW(1575) & ChrW(1605) & ChrW(1593) & ChrW(58)
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
