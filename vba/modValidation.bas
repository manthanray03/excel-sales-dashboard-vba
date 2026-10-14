Attribute VB_Name = "modValidation"
'===============================================================================
' modValidation
' Small rules engine. Rules live in tblRules (Validation_Rules sheet) so they can
' be changed without touching code. Every row in tblStaging is checked against
' every active rule:
'   - Error   -> the row is kept out of tblSales
'   - Warning -> the row is loaded but flagged
' All problems are written to tblExceptions with the rule and the bad value.
'
' Rule types:
'   Required      field must not be blank
'   Range         number between Param1 and Param2
'   WholeNumber   whole number between Param1 and Param2
'   DateRange     date between Param1 and Param2 (yyyy-mm-dd or TODAY)
'   Compare       field <op> other field   (Param1 = op, Param2 = other field)
'   MaxDaysAfter  field - Param1 field <= Param2 days
'   Lookup        value exists in table Param1, column Param2
'   List          value is one of Param1 (comma separated)
'   Unique        combination of fields (Field = "A+B") appears once
'   Tolerance     |field - Param1 field| / Param1 field <= Param2
'
' RunValidation reads/writes the tables; ValidateData does the actual work on
' arrays so it can be tested without a workbook.
'===============================================================================
Option Explicit

Public Type ValidationResult
    TotalRows As Long
    PassedRows As Long
    FailedRows As Long
    WarningRows As Long
    ExceptionCount As Long
    WarningCount As Long
End Type

Private Const STATUS_OK As Long = 0
Private Const STATUS_WARNING As Long = 1
Private Const STATUS_ERROR As Long = 2

' rowStatus() comes back 1-based, one entry per staging row (0 ok, 1 warning, 2 error)
Public Function RunValidation(ByRef rowStatus() As Long) As ValidationResult
    Dim res As ValidationResult
    Dim loStage As ListObject, loRules As ListObject, loExc As ListObject
    Dim data As Variant, rules As Variant, exc() As Variant, nExc As Long, nRows As Long, counts() As Long

    Set loStage = RequireTable(STAGING_TABLE)
    Set loRules = RequireTable("tblRules")
    Set loExc = RequireTable("tblExceptions")

    nRows = TableRowCount(loStage)
    If nRows = 0 Then
        WriteTable loExc, Empty, 0
        RunValidation = res
        Exit Function
    End If
    rules = TableToArray(loRules)
    If IsEmpty(rules) Then Err.Raise vbObjectError + 530, , "There are no validation rules in tblRules."
    data = TableToArray(loStage)

    ValidateData data, HeaderMap(loStage), nRows, rules, HeaderMap(loRules), _
                 LoadLookupSets(rules, HeaderMap(loRules)), Now, rowStatus, exc, nExc, counts
    WriteTable loExc, exc, nExc

    res.TotalRows = counts(0)
    res.PassedRows = counts(1)
    res.FailedRows = counts(2)
    res.WarningRows = counts(3)
    res.ExceptionCount = counts(4)
    res.WarningCount = counts(5)
    RunValidation = res
End Function

' For every active Lookup rule: the set of allowed values from table Param1, column Param2.
' Keyed "table|column" (lower case).
Public Function LoadLookupSets(ByRef rules As Variant, ByVal rMap As Collection) As Collection
    Dim sets As New Collection, i As Long, key As String, tbl As String, col As String
    Dim lo As ListObject, vals As Variant, cellValue As Variant, valueSet As Collection

    For i = 1 To UBound(rules, 1)
        If LCase$(ValueToText(rules(i, FieldIndex(rMap, "RuleType")))) = "lookup" Then
            tbl = ValueToText(rules(i, FieldIndex(rMap, "Param1")))
            col = ValueToText(rules(i, FieldIndex(rMap, "Param2")))
            key = LCase$(tbl & "|" & col)
            If Not KeyExists(sets, key) Then
                Set valueSet = New Collection
                Set lo = RequireTable(tbl)
                If TableRowCount(lo) > 0 Then
                    vals = lo.ListColumns(col).DataBodyRange.Value
                    If IsArray(vals) Then
                        For Each cellValue In vals
                            AddToSet valueSet, UCase$(ValueToText(cellValue))
                        Next cellValue
                    Else
                        AddToSet valueSet, UCase$(ValueToText(vals))
                    End If
                End If
                sets.Add valueSet, key
            End If
        End If
    Next i
    Set LoadLookupSets = sets
End Function

'------------------------------------------------------------------------------
' The engine. data = staging rows (1-based 2D array), map = column name -> index,
' rules = tblRules rows, lookupSets from LoadLookupSets. Fills rowStatus(),
' exc() (rows x 10, same columns as tblExceptions) and counts():
'   0 rows checked, 1 passed, 2 failed (errors), 3 rows with only warnings,
'   4 exceptions written, 5 warnings written
'------------------------------------------------------------------------------
Public Sub ValidateData(ByRef data As Variant, ByVal map As Collection, ByVal nRows As Long, _
                        ByRef rules As Variant, ByVal rMap As Collection, ByVal lookupSets As Collection, _
                        ByVal stamp As Date, ByRef rowStatus() As Long, ByRef exc() As Variant, _
                        ByRef nExc As Long, ByRef counts() As Long)
    Dim nRules As Long

    ' rule definitions (1..nRules)
    Dim rId() As String, rField() As String, rType() As String, rSev() As String, rDesc() As String
    Dim rP1() As Variant, rP2() As Variant, colA() As Long, colB() As Long
    Dim uniqCols() As Variant, valueSets() As Object, dMin() As Date, dMax() As Date

    Dim i As Long, r As Long, k As Long, worst As Long, failed As Boolean, skip As Boolean
    Dim v As Variant, w As Variant, shown As String, key As String, parts As Variant
    Dim cRaw As Long, cFile As Long, cOrder As Long, cProd As Long, nDef As Long

    ReDim counts(0 To 5)
    counts(0) = nRows
    nExc = 0
    ReDim exc(1 To nRows + 50, 1 To 10)
    If nRows = 0 Then Exit Sub
    ReDim rowStatus(1 To nRows)
    cRaw = FieldIndex(map, "RawRowID")
    cFile = FieldIndex(map, "SourceFile")
    cOrder = FieldIndex(map, "OrderID")
    cProd = FieldIndex(map, "ProductID")

    '--- active rules ------------------------------------------------------------
    nDef = UBound(rules, 1)
    ReDim rId(1 To nDef): ReDim rField(1 To nDef): ReDim rType(1 To nDef)
    ReDim rSev(1 To nDef): ReDim rDesc(1 To nDef)
    ReDim rP1(1 To nDef): ReDim rP2(1 To nDef)
    ReDim colA(1 To nDef): ReDim colB(1 To nDef)
    ReDim uniqCols(1 To nDef): ReDim valueSets(1 To nDef)
    ReDim dMin(1 To nDef): ReDim dMax(1 To nDef)

    For i = 1 To nDef
        If UCase$(ValueToText(rules(i, FieldIndex(rMap, "Active")))) = "Y" Then
            nRules = nRules + 1
            rId(nRules) = ValueToText(rules(i, FieldIndex(rMap, "RuleID")))
            rField(nRules) = ValueToText(rules(i, FieldIndex(rMap, "Field")))
            rType(nRules) = LCase$(ValueToText(rules(i, FieldIndex(rMap, "RuleType"))))
            rP1(nRules) = rules(i, FieldIndex(rMap, "Param1"))
            rP2(nRules) = rules(i, FieldIndex(rMap, "Param2"))
            rSev(nRules) = ValueToText(rules(i, FieldIndex(rMap, "Severity")))
            rDesc(nRules) = ValueToText(rules(i, FieldIndex(rMap, "Description")))
            PrepareRule nRules, rType(nRules), rField(nRules), rP1(nRules), rP2(nRules), map, lookupSets, _
                        colA, colB, uniqCols, valueSets, dMin, dMax
        End If
    Next i

    '--- check every row -----------------------------------------------------------
    For i = 1 To nRows
        worst = STATUS_OK
        For r = 1 To nRules
            failed = False
            skip = False
            shown = ""
            Select Case rType(r)
                Case "required"
                    v = data(i, colA(r))
                    failed = IsBlankValue(v)

                Case "range", "wholenumber"
                    v = data(i, colA(r))
                    shown = DisplayValue(v)
                    If Not IsNumberValue(v) Then
                        failed = True
                    Else
                        failed = (v < ParamNumber(rP1(r)) Or v > ParamNumber(rP2(r)))
                        If rType(r) = "wholenumber" And Not failed Then failed = (v <> Int(v))
                    End If

                Case "daterange"
                    v = data(i, colA(r))
                    If Not IsBlankValue(v) Then
                        w = ToDateValue(v)
                        shown = DisplayValue(v)
                        If IsEmpty(w) Then
                            failed = True
                        Else
                            failed = (w < dMin(r) Or w > dMax(r))
                        End If
                    End If

                Case "compare"
                    v = data(i, colA(r)): w = data(i, colB(r))
                    If Not IsBlankValue(v) And Not IsBlankValue(w) Then
                        failed = Not CompareValues(CDbl(v), CDbl(w), ValueToText(rP1(r)))
                        shown = DisplayValue(v) & " vs " & DisplayValue(w)
                    End If

                Case "maxdaysafter"
                    v = data(i, colA(r)): w = data(i, colB(r))
                    If Not IsBlankValue(v) And Not IsBlankValue(w) Then
                        failed = (CDbl(v) - CDbl(w) > ParamNumber(rP2(r)))
                        shown = Format$(CDbl(v) - CDbl(w), "0") & " days"
                    End If

                Case "lookup", "list"
                    v = data(i, colA(r))
                    failed = IsBlankValue(v)
                    If Not failed Then failed = Not KeyExists(valueSets(r), UCase$(ValueToText(v)))

                Case "unique"
                    parts = uniqCols(r)
                    key = ""
                    For k = LBound(parts) To UBound(parts)
                        v = data(i, parts(k))
                        If IsBlankValue(v) Then skip = True
                        If k > LBound(parts) Then key = key & "|"
                        key = key & UCase$(DisplayValue(v))
                    Next k
                    If Not skip Then
                        shown = key
                        If KeyExists(valueSets(r), key) Then
                            failed = True
                        Else
                            valueSets(r).Add data(i, cRaw), key
                        End If
                    End If

                Case "tolerance"
                    v = data(i, colA(r)): w = data(i, colB(r))
                    If IsNumberValue(v) And IsNumberValue(w) Then
                        If w <> 0 Then
                            failed = (Abs(v - w) / w > ParamNumber(rP2(r)))
                            shown = DisplayValue(v) & " vs list " & DisplayValue(w)
                        End If
                    End If
            End Select

            If failed Then
                If Len(shown) = 0 And rType(r) <> "unique" Then shown = DisplayValue(data(i, colA(r)))
                nExc = nExc + 1
                If nExc > UBound(exc, 1) Then GrowRows exc, UBound(exc, 1) * 2
                exc(nExc, 1) = data(i, cRaw)
                exc(nExc, 2) = data(i, cFile)
                exc(nExc, 3) = data(i, cOrder)
                exc(nExc, 4) = data(i, cProd)
                exc(nExc, 5) = rId(r)
                exc(nExc, 6) = rSev(r)
                exc(nExc, 7) = rField(r)
                exc(nExc, 8) = rDesc(r)
                exc(nExc, 9) = shown
                exc(nExc, 10) = stamp
                If LCase$(rSev(r)) = "error" Then
                    worst = STATUS_ERROR
                Else
                    If worst < STATUS_WARNING Then worst = STATUS_WARNING
                    counts(5) = counts(5) + 1
                End If
            End If
        Next r

        rowStatus(i) = worst
        Select Case worst
            Case STATUS_ERROR: counts(2) = counts(2) + 1
            Case STATUS_WARNING: counts(3) = counts(3) + 1
        End Select
    Next i

    counts(1) = nRows - counts(2)
    counts(4) = nExc
End Sub

' Works out the column numbers / value sets a rule needs, once, before the row loop
Private Sub PrepareRule(ByVal r As Long, ByVal ruleType As String, ByVal fieldName As String, _
                        ByVal p1 As Variant, ByVal p2 As Variant, ByVal map As Collection, _
                        ByVal lookupSets As Collection, ByRef colA() As Long, ByRef colB() As Long, _
                        ByRef uniqCols() As Variant, ByRef valueSets() As Object, _
                        ByRef dMin() As Date, ByRef dMax() As Date)
    Dim fieldList As Variant, cols() As Long, k As Long, key As String

    Select Case ruleType
        Case "unique"
            fieldList = Split(fieldName, "+")
            ReDim cols(0 To UBound(fieldList))
            For k = 0 To UBound(fieldList)
                cols(k) = FieldIndex(map, Trim$(fieldList(k)))
            Next k
            uniqCols(r) = cols
            Set valueSets(r) = New Collection

        Case "compare", "tolerance", "maxdaysafter"
            colA(r) = FieldIndex(map, fieldName)
            If ruleType = "compare" Then
                colB(r) = FieldIndex(map, ValueToText(p2))
            Else
                colB(r) = FieldIndex(map, ValueToText(p1))
            End If

        Case "lookup"
            colA(r) = FieldIndex(map, fieldName)
            key = LCase$(ValueToText(p1) & "|" & ValueToText(p2))
            If Not KeyExists(lookupSets, key) Then Err.Raise vbObjectError + 533, , "No lookup values loaded for " & key & "."
            Set valueSets(r) = lookupSets.Item(key)

        Case "list"
            colA(r) = FieldIndex(map, fieldName)
            Set valueSets(r) = New Collection
            fieldList = Split(ValueToText(p1), ",")
            For k = 0 To UBound(fieldList)
                AddToSet valueSets(r), UCase$(Trim$(fieldList(k)))
            Next k

        Case "daterange"
            colA(r) = FieldIndex(map, fieldName)
            dMin(r) = ParamDate(p1)
            dMax(r) = ParamDate(p2)

        Case "required", "range", "wholenumber"
            colA(r) = FieldIndex(map, fieldName)

        Case Else
            Err.Raise vbObjectError + 531, , "Unknown rule type '" & ruleType & "' in tblRules."
    End Select
End Sub

Private Sub AddToSet(ByVal col As Collection, ByVal key As String)
    If Len(key) = 0 Then Exit Sub
    If Not KeyExists(col, key) Then col.Add True, key
End Sub

' Params can be typed as numbers or text in the rules table
Private Function ParamNumber(ByVal p As Variant) As Double
    If IsNumberValue(p) Then
        ParamNumber = CDbl(p)
    Else
        ParamNumber = Val(ValueToText(p))
    End If
End Function

Private Function ParamDate(ByVal p As Variant) As Date
    Dim d As Variant
    If UCase$(ValueToText(p)) = "TODAY" Then
        ParamDate = Date
        Exit Function
    End If
    d = ToDateValue(p)
    If IsEmpty(d) Then Err.Raise vbObjectError + 532, , "Bad date parameter '" & ValueToText(p) & "' in tblRules."
    ParamDate = d
End Function
