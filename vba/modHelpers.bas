Attribute VB_Name = "modHelpers"
'===============================================================================
' modHelpers
' Pure helper functions (no worksheet access) - CSV parsing, text cleanup,
' number/date parsing, sorting and a few calculations. Keeping these free of
' Excel objects makes them easy to test on their own.
'===============================================================================
Option Explicit

'------------------------------------------------------------------------------
' CSV
'------------------------------------------------------------------------------

' Splits one CSV line into fields. Handles quoted fields ("$1,234.50") and
' doubled quotes inside quotes. Returns a 0-based String array.
Public Function SplitCsvLine(ByVal lineText As String) As Variant
    Dim fields() As String, n As Long, i As Long
    Dim ch As String, cur As String, inQuotes As Boolean

    If InStr(lineText, """") = 0 Then
        SplitCsvLine = Split(lineText, ",")
        Exit Function
    End If

    ReDim fields(0 To 31)
    i = 1
    Do While i <= Len(lineText)
        ch = Mid$(lineText, i, 1)
        If inQuotes Then
            If ch = """" Then
                If Mid$(lineText, i + 1, 1) = """" Then
                    cur = cur & """"
                    i = i + 1
                Else
                    inQuotes = False
                End If
            Else
                cur = cur & ch
            End If
        Else
            Select Case ch
                Case """"
                    inQuotes = True
                Case ","
                    If n > UBound(fields) Then
                        ReDim Preserve fields(0 To n * 2)
                    End If
                    fields(n) = cur
                    n = n + 1
                    cur = ""
                Case Else
                    cur = cur & ch
            End Select
        End If
        i = i + 1
    Loop

    If n > UBound(fields) Then
        ReDim Preserve fields(0 To n * 2)
    End If
    fields(n) = cur
    ReDim Preserve fields(0 To n)
    SplitCsvLine = fields
End Function

Public Function QuoteCount(ByVal s As String) As Long
    QuoteCount = Len(s) - Len(Replace(s, """", ""))
End Function

Public Function AllFieldsBlank(ByVal fields As Variant) As Boolean
    Dim i As Long
    For i = LBound(fields) To UBound(fields)
        If Len(Trim$(fields(i))) > 0 Then Exit Function
    Next i
    AllFieldsBlank = True
End Function

'------------------------------------------------------------------------------
' Text cleanup (same rules as the Power Query step)
'------------------------------------------------------------------------------

' "  john    smith " -> "john smith"
Public Function CleanSpaces(ByVal t As String) As String
    t = Trim$(t)
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    CleanSpaces = t
End Function

' Capitalises the first letter of each word. Not using StrConv(vbProperCase)
' because it turns "St. John's" into "St. John'S".
Public Function ProperCase(ByVal t As String) As String
    Dim words() As String, i As Long
    t = LCase$(CleanSpaces(t))
    If Len(t) = 0 Then Exit Function
    words = Split(t, " ")
    For i = 0 To UBound(words)
        words(i) = UCase$(Left$(words(i), 1)) & Mid$(words(i), 2)
    Next i
    ProperCase = Join(words, " ")
End Function

' Key used to match province spellings: "B.C." -> "BC", "Nova Scotia" -> "NOVASCOTIA"
Public Function NormalizeKey(ByVal t As String) As String
    t = Replace(t, " ", "")
    t = Replace(t, ".", "")
    t = Replace(t, "-", "")
    t = Replace(t, "'", "")
    NormalizeKey = UCase$(t)
End Function

' "In Store", "instore", "WEB"... -> "In-Store" / "Online" (blank if unknown)
Public Function StandardChannel(ByVal t As String) As String
    Dim k As String
    k = LCase$(Replace(Replace(Replace(t, " ", ""), "-", ""), "_", ""))
    Select Case k
        Case "instore", "store", "pos"
            StandardChannel = "In-Store"
        Case "online", "web", "ecommerce"
            StandardChannel = "Online"
        Case Else
            StandardChannel = ""
    End Select
End Function

' Cell value -> trimmed text (dates come back as yyyy-mm-dd)
Public Function ValueToText(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If VarType(v) = vbDate Then
        ValueToText = Format$(v, "yyyy-mm-dd")
    Else
        ValueToText = Trim$(CStr(v))
    End If
End Function

'------------------------------------------------------------------------------
' Number / date parsing. Val() is used on purpose - it always reads "." as the
' decimal point, so the results don't change with the PC's regional settings.
'------------------------------------------------------------------------------

' True for things like 12, -3, 0.15, .5 (no thousands separators, no letters)
Public Function IsPlainNumber(ByVal s As String) As Boolean
    Dim i As Long, ch As String, digits As Long, dots As Long
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then
            digits = digits + 1
        ElseIf ch = "." Then
            dots = dots + 1
            If dots > 1 Then Exit Function
        ElseIf (ch = "-" Or ch = "+") And i = 1 Then
            ' leading sign is fine
        Else
            Exit Function
        End If
    Next i
    IsPlainNumber = (digits > 0)
End Function

Public Function IsDigitsOnly(ByVal s As String) As Boolean
    Dim i As Long, ch As String
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    IsDigitsOnly = True
End Function

' "$1,234.50" / " 3 " / "2.0" -> number. Returns Empty when it isn't a number.
Public Function ParseNumberText(ByVal t As String) As Variant
    Dim s As String
    s = Replace(Replace(Replace(t, "$", ""), ",", ""), " ", "")
    If IsPlainNumber(s) Then
        ParseNumberText = Val(s)
    Else
        ParseNumberText = Empty
    End If
End Function

' "10%" -> 0.1, "0.1" -> 0.1, "" -> 0. Returns Empty for junk.
Public Function ParseDiscountText(ByVal t As String) As Variant
    Dim s As String
    s = Replace(t, " ", "")
    If Len(s) = 0 Then
        ParseDiscountText = 0#
    ElseIf Right$(s, 1) = "%" Then
        s = Left$(s, Len(s) - 1)
        If IsPlainNumber(s) Then ParseDiscountText = Val(s) / 100 Else ParseDiscountText = Empty
    ElseIf IsPlainNumber(s) Then
        ParseDiscountText = Val(s)
    Else
        ParseDiscountText = Empty
    End If
End Function

' Exports use YYYY-MM-DD, a few store terminals send MM/DD/YYYY.
' Returns a Date, or Empty when the text isn't a real date (e.g. 2024-02-30).
Public Function ParseExportDate(ByVal t As String) As Variant
    Dim s As String, parts() As String, isUS As Boolean
    Dim y As Long, m As Long, d As Long

    ParseExportDate = Empty
    s = Trim$(t)
    isUS = (InStr(s, "/") > 0)
    If isUS Then
        parts = Split(s, "/")
    ElseIf InStr(s, "-") > 0 Then
        parts = Split(Left$(s, 10), "-")
    Else
        Exit Function
    End If
    If UBound(parts) <> 2 Then Exit Function
    If Not (IsDigitsOnly(parts(0)) And IsDigitsOnly(parts(1)) And IsDigitsOnly(parts(2))) Then Exit Function

    If isUS Then
        m = CLng(parts(0)): d = CLng(parts(1)): y = CLng(parts(2))
    Else
        y = CLng(parts(0)): m = CLng(parts(1)): d = CLng(parts(2))
    End If
    If y < 1900 Or y > 9999 Or m < 1 Or m > 12 Or d < 1 Then Exit Function
    If d > Day(DateSerial(y, m + 1, 0)) Then Exit Function

    ParseExportDate = DateSerial(y, m, d)
End Function

'------------------------------------------------------------------------------
' Value checks
'------------------------------------------------------------------------------
Public Function IsBlankValue(ByVal v As Variant) As Boolean
    If IsError(v) Then
        IsBlankValue = True
    ElseIf IsEmpty(v) Or IsNull(v) Then
        IsBlankValue = True
    ElseIf VarType(v) = vbString Then
        IsBlankValue = (Len(Trim$(v)) = 0)
    End If
End Function

' IsNumeric() says True for "12" and for Empty, which is not what we want here
Public Function IsNumberValue(ByVal v As Variant) As Boolean
    Select Case VarType(v)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte
            IsNumberValue = True
    End Select
End Function

' Date cell -> Date, Empty if it isn't one
Public Function ToDateValue(ByVal v As Variant) As Variant
    ToDateValue = Empty
    If VarType(v) = vbDate Then
        ToDateValue = v
    ElseIf IsNumberValue(v) Then
        ToDateValue = CDate(v)
    ElseIf VarType(v) = vbString Then
        ToDateValue = ParseExportDate(CStr(v))
    End If
End Function

' Text shown in the Exceptions sheet for a value
Public Function DisplayValue(ByVal v As Variant) As String
    If IsBlankValue(v) Then Exit Function
    If VarType(v) = vbDate Then
        DisplayValue = Format$(v, "yyyy-mm-dd")
    ElseIf IsNumberValue(v) Then
        If v = Int(v) Then DisplayValue = Format$(v, "0") Else DisplayValue = CStr(v)
    Else
        DisplayValue = CStr(v)
    End If
End Function

Public Function CompareValues(ByVal a As Double, ByVal b As Double, ByVal op As String) As Boolean
    Select Case Trim$(op)
        Case ">=": CompareValues = (a >= b)
        Case "<=": CompareValues = (a <= b)
        Case ">":  CompareValues = (a > b)
        Case "<":  CompareValues = (a < b)
        Case "=":  CompareValues = (a = b)
        Case "<>": CompareValues = (a <> b)
        Case Else: Err.Raise vbObjectError + 600, , "Unknown comparison operator '" & op & "'"
    End Select
End Function

'------------------------------------------------------------------------------
' Collections (used instead of Scripting.Dictionary so this also runs on a Mac)
'------------------------------------------------------------------------------
Public Function KeyExists(ByVal col As Collection, ByVal key As String) As Boolean
    Dim isObj As Boolean
    On Error GoTo NotFound
    isObj = IsObject(col.Item(key))     ' works for both object and plain items
    KeyExists = True
    Exit Function
NotFound:
    KeyExists = False
End Function

' Column name -> column number (case-insensitive) from a list of headers
Public Function MapFromHeaders(ByVal headers As Variant) As Collection
    Dim col As New Collection, i As Long, k As Long
    For i = LBound(headers) To UBound(headers)
        k = k + 1
        If Not KeyExists(col, LCase$(CStr(headers(i)))) Then col.Add k, LCase$(CStr(headers(i)))
    Next i
    Set MapFromHeaders = col
End Function

' Same as modUtils.ColIndex but without needing a table (used by the pure functions)
Public Function FieldIndex(ByVal map As Collection, ByVal columnName As String) As Long
    If KeyExists(map, LCase$(columnName)) Then
        FieldIndex = map.Item(LCase$(columnName))
    Else
        Err.Raise vbObjectError + 514, "FieldIndex", "Column '" & columnName & "' not found."
    End If
End Function

'------------------------------------------------------------------------------
' Math
'------------------------------------------------------------------------------

' Round to 2 decimals, halves away from zero (VBA's Round() does banker's rounding)
Public Function Round2(ByVal x As Double) As Double
    If x >= 0 Then
        Round2 = Int(x * 100 + 0.5) / 100
    Else
        Round2 = -Int(-x * 100 + 0.5) / 100
    End If
End Function

'------------------------------------------------------------------------------
' Sorting
'------------------------------------------------------------------------------

' Sorts idx() so that keys(idx(i)) is ascending; ties are broken by tieKeys()
' (text, ascending). keys/tieKeys are not moved - only the index array is.
' Pass negative keys to sort descending.
Public Sub QuickSortIndex(ByRef keys() As Double, ByRef tieKeys() As String, ByRef idx() As Long, _
                          ByVal first As Long, ByVal last As Long)
    Dim i As Long, j As Long, pivot As Long, tmp As Long
    If first >= last Then Exit Sub
    pivot = idx((first + last) \ 2)
    i = first
    j = last
    Do While i <= j
        Do While IndexLess(keys, tieKeys, idx(i), pivot)
            i = i + 1
        Loop
        Do While IndexLess(keys, tieKeys, pivot, idx(j))
            j = j - 1
        Loop
        If i <= j Then
            tmp = idx(i): idx(i) = idx(j): idx(j) = tmp
            i = i + 1
            j = j - 1
        End If
    Loop
    If first < j Then QuickSortIndex keys, tieKeys, idx, first, j
    If i < last Then QuickSortIndex keys, tieKeys, idx, i, last
End Sub

Private Function IndexLess(ByRef keys() As Double, ByRef tieKeys() As String, ByVal a As Long, ByVal b As Long) As Boolean
    If keys(a) < keys(b) Then
        IndexLess = True
    ElseIf keys(a) = keys(b) Then
        IndexLess = (StrComp(tieKeys(a), tieKeys(b), vbBinaryCompare) < 0)
    End If
End Function

' Sorts a 1-based Double array ascending (used for the RFM quintiles)
Public Sub QuickSortDoubles(ByRef arr() As Double, ByVal first As Long, ByVal last As Long)
    Dim i As Long, j As Long, pivot As Double, tmp As Double
    If first >= last Then Exit Sub
    pivot = arr((first + last) \ 2)
    i = first
    j = last
    Do While i <= j
        Do While arr(i) < pivot
            i = i + 1
        Loop
        Do While pivot < arr(j)
            j = j - 1
        Loop
        If i <= j Then
            tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            i = i + 1
            j = j - 1
        End If
    Loop
    If first < j Then QuickSortDoubles arr, first, j
    If i < last Then QuickSortDoubles arr, i, last
End Sub

' Number of values in a sorted array that are strictly less than x
Public Function CountLess(ByRef sorted() As Double, ByVal n As Long, ByVal x As Double) As Long
    Dim lo As Long, hi As Long, half As Long
    lo = 1: hi = n + 1
    Do While lo < hi
        half = (lo + hi) \ 2
        If sorted(half) < x Then lo = half + 1 Else hi = half
    Loop
    CountLess = lo - 1
End Function

' Number of values in a sorted array that are strictly greater than x
Public Function CountGreater(ByRef sorted() As Double, ByVal n As Long, ByVal x As Double) As Long
    Dim lo As Long, hi As Long, half As Long
    lo = 1: hi = n + 1
    Do While lo < hi
        half = (lo + hi) \ 2
        If sorted(half) <= x Then lo = half + 1 Else hi = half
    Loop
    CountGreater = n - (lo - 1)
End Function

'------------------------------------------------------------------------------
' Business calculations
'------------------------------------------------------------------------------

' 1-5 score where a higher value is better: share of customers strictly below
Public Function QuintileScore(ByVal countBelow As Long, ByVal n As Long) As Long
    Dim s As Long
    If n <= 0 Then QuintileScore = 1: Exit Function
    s = 1 + Int(5 * countBelow / n)
    If s > 5 Then s = 5
    QuintileScore = s
End Function

Public Function RfmSegment(ByVal r As Long, ByVal f As Long, ByVal m As Long) As String
    If r >= 4 And f >= 4 And m >= 4 Then
        RfmSegment = "Champions"
    ElseIf r >= 3 And f >= 4 Then
        RfmSegment = "Loyal Customers"
    ElseIf r >= 4 And f = 1 Then
        RfmSegment = "New Customers"
    ElseIf r >= 3 And f >= 2 Then
        RfmSegment = "Potential Loyalists"
    ElseIf r <= 2 And f >= 3 Then
        RfmSegment = "At Risk"
    ElseIf r <= 2 Then
        RfmSegment = "Hibernating"
    Else
        RfmSegment = "Needs Attention"
    End If
End Function

' Marginal (stepped) commission. tierFrom() holds the attainment where each
' tier starts (0, 0.8, 1.0, 1.2...) and tierRate() the rate paid on revenue
' inside that tier. Arrays are 1-based and sorted by tierFrom.
Public Function TieredCommission(ByVal revenue As Double, ByVal quota As Double, _
                                 ByRef tierFrom() As Double, ByRef tierRate() As Double, _
                                 ByVal nTiers As Long) As Double
    Dim i As Long, lower As Double, upper As Double, top As Double, total As Double
    For i = 1 To nTiers
        lower = tierFrom(i) * quota
        If i < nTiers Then
            upper = tierFrom(i + 1) * quota
            If revenue < upper Then top = revenue Else top = upper
        Else
            top = revenue
        End If
        If top > lower Then total = total + (top - lower) * tierRate(i)
    Next i
    TieredCommission = Round2(total)
End Function

'------------------------------------------------------------------------------
' Fallback cleaning (used only if Power Query can't run). Takes one raw row
' (16 text fields in export order) and returns the 24 staging columns.
' Lookups: aliasMap  normalized alias -> province code
'          provMap   province code -> row in provInfo (1 = name, 2 = region)
'          prodMap   product id -> row in prodInfo (1 = name, 2 = category,
'                    3 = subcategory, 4 = list price, 5 = unit cost)
'------------------------------------------------------------------------------
Public Function CleanRawRow(ByVal rawRowId As Variant, ByVal sourceFile As String, ByRef f() As String, _
                            ByVal aliasMap As Collection, ByVal provMap As Collection, ByRef provInfo() As Variant, _
                            ByVal prodMap As Collection, ByRef prodInfo() As Variant, _
                            ByRef provinceKey As String) As Variant
    Dim out(1 To 24) As Variant, code As String, pid As String, k As Long

    out(1) = rawRowId
    out(2) = sourceFile
    out(3) = UCase$(f(0))                         ' OrderID
    out(4) = ParseExportDate(f(1))                ' OrderDate
    out(5) = ParseExportDate(f(2))                ' ShipDate
    out(6) = UCase$(f(3))                         ' CustomerID
    out(7) = ProperCase(f(4))                     ' CustomerName
    out(8) = ProperCase(f(5))                     ' Segment
    out(9) = ProperCase(f(6))                     ' City

    provinceKey = NormalizeKey(f(7))
    If KeyExists(aliasMap, provinceKey) Then
        code = aliasMap.Item(provinceKey)
        out(10) = code
        If KeyExists(provMap, code) Then
            k = provMap.Item(code)
            out(11) = provInfo(k, 1)
            out(12) = provInfo(k, 2)
        End If
    End If

    pid = UCase$(f(8))
    out(13) = pid
    If KeyExists(prodMap, pid) Then
        k = prodMap.Item(pid)
        out(14) = prodInfo(k, 1)
        out(15) = prodInfo(k, 2)
        out(16) = prodInfo(k, 3)
        out(19) = prodInfo(k, 4)
        out(20) = prodInfo(k, 5)
    Else
        out(14) = f(9)
    End If

    out(17) = ParseNumberText(f(10))              ' Quantity
    out(18) = ParseNumberText(f(11))              ' UnitPrice
    out(21) = ParseDiscountText(f(12))            ' Discount
    out(22) = StandardChannel(f(13))              ' Channel
    If Len(out(22)) = 0 Then out(22) = Empty
    out(23) = CleanSpaces(f(14))                  ' SalesRep
    out(24) = CleanSpaces(f(15))                  ' PaymentMethod

    CleanRawRow = out
End Function

' Key used to drop exact duplicate rows (same columns Power Query compares)
Public Function DuplicateKey(ByRef stagingRow As Variant, ByVal provinceKey As String) As String
    Dim parts(1 To 15) As String
    parts(1) = KeyPart(stagingRow(3)):  parts(2) = KeyPart(stagingRow(4)):  parts(3) = KeyPart(stagingRow(5))
    parts(4) = KeyPart(stagingRow(6)):  parts(5) = KeyPart(stagingRow(7)):  parts(6) = KeyPart(stagingRow(8))
    parts(7) = KeyPart(stagingRow(9)):  parts(8) = provinceKey:      parts(9) = KeyPart(stagingRow(13))
    parts(10) = KeyPart(stagingRow(17)): parts(11) = KeyPart(stagingRow(18)): parts(12) = KeyPart(stagingRow(21))
    parts(13) = KeyPart(stagingRow(22)): parts(14) = KeyPart(stagingRow(23)): parts(15) = KeyPart(stagingRow(24))
    DuplicateKey = Join(parts, Chr$(30))
End Function

Private Function KeyPart(ByVal v As Variant) As String
    If IsEmpty(v) Then
        KeyPart = Chr$(31)
    ElseIf VarType(v) = vbDate Then
        KeyPart = Format$(v, "yyyy-mm-dd")
    Else
        KeyPart = CStr(v)
    End If
End Function
