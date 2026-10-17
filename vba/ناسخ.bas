Attribute VB_Name = "Nasikh"
'============================================================
'  ناسخ — ماكرو Word للبحث في مكتبة الجامع وإدراج النص
'          مع المرجع  <اسم الكتاب الجزء/الصفحة>  داخل المستند.
'
'  يتصل بخادم الجامع العام (MCP) عبر HTTPS، بلا مفتاح.
'  يحتاج اتصالًا بالإنترنت.
'
'  الاستخدام: شغّل الماكرو  NasikhSearch  (أو من الزر في الشريط).
'============================================================
Option Explicit

Private Const MCP_URL As String = "https://aljam3.com/mcp"


'------------------------------------------------------------
' نقطة الدخول: يسأل عن العبارة، يبحث، ويُدرج النتائج.
'------------------------------------------------------------
Public Sub NasikhSearch()
    Dim query As String
    query = InputBox("اكتب عبارة البحث في مكتبة الجامع:", "ناسخ — بحث")
    If Trim(query) = "" Then Exit Sub

    Dim countStr As String
    Dim response As String
    On Error GoTo HttpError
    response = McpSearch(query, 10)
    On Error GoTo 0

    If Len(response) = 0 Then
        MsgBox "لم تصل نتائج. تأكد من الاتصال بالإنترنت.", vbExclamation, "ناسخ"
        Exit Sub
    End If

    InsertResultsFromJson response, query
    Exit Sub

HttpError:
    MsgBox "تعذّر الاتصال بالجامع:" & vbCrLf & Err.Description, vbCritical, "ناسخ"
End Sub


'------------------------------------------------------------
' يُرسل طلب MCP (JSON-RPC) إلى أداة search ويعيد نص الرد.
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
' يحلّل رد الجامع ويُدرج كل نتيجة (المقتطف + المرجع) في المستند.
' ملاحظة: محلّل مبسّط يستخرج الحقول المطلوبة فقط (book title/number/excerpt).
'------------------------------------------------------------
Private Sub InsertResultsFromJson(ByVal json As String, ByVal query As String)
    ' لو كان الرد تدفّق أحداث (SSE) فخذ ما بعد أول "data:"
    Dim p As Long
    p = InStr(json, "data:")
    If p > 0 Then json = Mid$(json, p + 5)

    Dim sel As Object
    Set sel = Selection

    ' عنوان
    sel.TypeText "نتائج البحث عن: «" & query & "»"
    sel.TypeParagraph
    sel.TypeParagraph

    ' مرّ على النتائج: نقسّم عند كل "excerpt" (أو snip) كبداية عنصر
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

        If bookName <> "" Or excerpt <> "" Then
            n = n + 1
            ' النص
            sel.TypeText excerpt
            sel.TypeParagraph
            ' المرجع  <اسم الكتاب الجزء/الصفحة>  محاذى لليمين
            Dim ref As String
            ref = BuildReference(bookName, vol, pageNum)
            InsertReferenceRight sel, ref
            sel.TypeParagraph
            sel.TypeParagraph
        End If
    Next i

    If n = 0 Then
        sel.TypeText "لا توجد نتائج مطابقة."
        sel.TypeParagraph
    End If
End Sub


'------------------------------------------------------------
' يبني نص المرجع:  <اسم الكتاب الجزء/الصفحة>
' يتجاهل الجزء إن لم يكن رقمًا نظيفًا، ويكتفي بالصفحة.
'------------------------------------------------------------
Private Function BuildReference(ByVal name As String, ByVal vol As String, ByVal page As String) As String
    If name = "" Then name = "كتاب غير معروف"
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
' يُدرج المرجع في سطر محاذى لليمين (مع علامة RLM).
'------------------------------------------------------------
Private Sub InsertReferenceRight(ByVal sel As Object, ByVal ref As String)
    sel.ParagraphFormat.Alignment = wdAlignParagraphRight
    sel.Font.Bold = True
    sel.TypeText ChrW(&H200F) & ref   ' RLM لضمان اتجاه القوس < >
    sel.Font.Bold = False
    sel.ParagraphFormat.Alignment = wdAlignParagraphRight
End Sub


'============================================================
'  أدوات مساعدة (محلّل JSON مبسّط + تنظيف نص)
'============================================================

' يقسّم الرد إلى عناصر عند كل بداية نتيجة (نبحث عن "excerpt" أو "snip").
Private Function SplitResults(ByVal json As String) As String()
    Dim marker As String
    marker = """excerpt"""
    If InStr(json, marker) = 0 Then marker = """snip"""
    If InStr(json, marker) = 0 Then marker = """content"""

    Dim parts() As String
    parts = Split(json, marker)

    ' أعِد لصق الواسم في بداية كل جزء (ما عدا الأول قبل أول نتيجة)
    Dim res() As String
    Dim cnt As Long
    cnt = UBound(parts) - LBound(parts)   ' عدد النتائج تقريبًا
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


' يستخرج قيمة مفتاح بسيط ("key":"value" أو "key":value) من مقطع JSON.
Private Function GetJsonValue(ByVal json As String, ByVal key As String) As String
    Dim pat As String
    pat = """" & key & """"
    Dim i As Long
    i = InStr(json, pat)
    If i = 0 Then Exit Function

    i = InStr(i + Len(pat), json, ":")
    If i = 0 Then Exit Function
    i = i + 1

    ' تخطّ الفراغات
    Do While i <= Len(json) And Mid$(json, i, 1) = " "
        i = i + 1
    Loop

    If i > Len(json) Then Exit Function

    Dim ch As String
    ch = Mid$(json, i, 1)
    If ch = """" Then
        ' قيمة نصية: اقرأ حتى علامة الاقتباس التالية غير المهرّبة
        Dim j As Long
        j = i + 1
        Dim out As String
        Do While j <= Len(json)
            Dim c As String
            c = Mid$(json, j, 1)
            If c = "\" Then
                ' حرف مهرّب
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
                        Dim hex As String
                        hex = Mid$(json, j + 2, 4)
                        out = out & ChrW(CLng("&H" & hex))
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
        ' قيمة رقمية/منطقية: اقرأ حتى , أو } أو ]
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


' يستخرج اسم الكتاب: قد يكون "book":{"name":"..."} أو "book_name":"..." أو "title":"..."
Private Function GetBookName(ByVal json As String) As String
    Dim i As Long
    i = InStr(json, """book""")
    If i > 0 Then
        ' خذ المقطع بعد "book": واستخرج name منه
        Dim sub1 As String
        sub1 = Mid$(json, i)
        GetBookName = GetJsonValue(sub1, "name")
        If GetBookName <> "" Then Exit Function
        GetBookName = GetJsonValue(sub1, "title")
        If GetBookName <> "" Then Exit Function
    End If
    GetBookName = GetJsonValue(json, "book_name")
    If GetBookName <> "" Then Exit Function
    GetBookName = GetJsonValue(json, "title")
End Function


' تنظيف النص: إزالة وسوم HTML وبقايا شائعة.
Private Function CleanText(ByVal s As String) As String
    If s = "" Then Exit Function
    ' أزل الوسوم <...>
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
    ' فكّ كيانات شائعة
    out = Replace(out, "&amp;", "&")
    out = Replace(out, "&lt;", "<")
    out = Replace(out, "&gt;", ">")
    out = Replace(out, "&quot;", """")
    ' قلّص الفراغات
    Do While InStr(out, "  ") > 0
        out = Replace(out, "  ", " ")
    Loop
    CleanText = Trim$(out)
End Function


' هل السلسلة أرقام فقط (عربية أو لاتينية)؟
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


' تهريب النص لإدراجه داخل JSON.
Private Function JsonEscape(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    JsonEscape = s
End Function
