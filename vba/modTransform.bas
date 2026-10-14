Attribute VB_Name = "modTransform"
'===============================================================================
' modTransform  -  "T" of the ETL
' The cleaning itself is done by the Power Query query "SalesStaging" (M code
' in the power-query folder of the repo, and a copy on the Settings sheet).
' This module makes sure the query exists and is loaded to tblStaging, then
' refreshes it.
'
' If Power Query can't run on the machine (older Excel, some Mac builds) the
' user can choose to run TransformWithVBA instead, which applies the same rules.
'===============================================================================
Option Explicit

Public Const STAGING_QUERY As String = "SalesStaging"
Public Const STAGING_TABLE As String = "tblStaging"

' Returns False if the step failed and the user chose not to use the fallback
Public Function RunTransform(ByRef methodUsed As String) As Boolean
    Dim errMsg As String

    If TryPowerQuery(errMsg) Then
        methodUsed = "Power Query"
        RunTransform = True
        Exit Function
    End If

    Application.ScreenUpdating = True
    If MsgBox("The Power Query step couldn't run on this computer:" & vbLf & vbLf & errMsg & vbLf & vbLf & _
              "Use the built-in VBA cleaning instead? (same rules, just slower)", _
              vbExclamation + vbYesNo, APP_TITLE) = vbYes Then
        Application.ScreenUpdating = False
        TransformWithVBA
        methodUsed = "VBA fallback"
        RunTransform = True
    Else
        methodUsed = "Power Query (failed)"
    End If
End Function

Private Function TryPowerQuery(ByRef errMsg As String) As Boolean
    On Error GoTo Failed
    EnsureStagingQuery
    RefreshStagingQuery
    TryPowerQuery = True
    Exit Function
Failed:
    errMsg = Err.Description
End Function

'------------------------------------------------------------------------------
' Power Query
'------------------------------------------------------------------------------

' Creates the query from the M code on the Settings sheet if it isn't in the
' workbook yet, and loads it to tblStaging on the Staging sheet.
Public Sub EnsureStagingQuery()
    Dim ws As Worksheet, lo As ListObject, mCode As String

    If Not QueryExists(STAGING_QUERY) Then
        mCode = CStr(NamedValue("PQ_StagingM"))
        If Len(mCode) = 0 Then Err.Raise vbObjectError + 520, , "The M code for " & STAGING_QUERY & " is missing (Settings sheet)."
        ThisWorkbook.Queries.Add Name:=STAGING_QUERY, Formula:=mCode, _
            Description:="Cleans tblRaw and enriches it with product + province lookups"
    End If

    Set lo = GetTable(STAGING_TABLE)
    If Not lo Is Nothing Then
        If lo.SourceType = xlSrcQuery Then Exit Sub   ' already wired up
        lo.Delete                                     ' plain copy from before the query existed
    End If

    Set ws = ThisWorkbook.Worksheets(SH_STAGING)
    Set lo = ws.ListObjects.Add(SourceType:=xlSrcExternal, _
        Source:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & STAGING_QUERY & _
                ";Extended Properties=""""", _
        Destination:=ws.Range("A3"))
    With lo.QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [" & STAGING_QUERY & "]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .Refresh BackgroundQuery:=False
    End With
    lo.DisplayName = STAGING_TABLE
    lo.TableStyle = "TableStyleLight9"
    FormatColumn lo, "OrderDate", "yyyy-mm-dd"
    FormatColumn lo, "ShipDate", "yyyy-mm-dd"
End Sub

Public Sub RefreshStagingQuery()
    Dim lo As ListObject
    Set lo = RequireTable(STAGING_TABLE)
    lo.QueryTable.Refresh BackgroundQuery:=False
End Sub

Private Function QueryExists(ByVal queryName As String) As Boolean
    Dim q As Object
    For Each q In ThisWorkbook.Queries
        If StrComp(q.Name, queryName, vbTextCompare) = 0 Then
            QueryExists = True
            Exit Function
        End If
    Next q
End Function

' Handy if the M code on the Settings sheet was edited: pushes it into the query
Public Sub UpdateQueryFromSettings()
    Dim mCode As String
    mCode = CStr(NamedValue("PQ_StagingM"))
    If QueryExists(STAGING_QUERY) Then
        ThisWorkbook.Queries(STAGING_QUERY).Formula = mCode
        MsgBox "Query updated. Run Full Refresh to use it.", vbInformation, APP_TITLE
    Else
        EnsureStagingQuery
    End If
End Sub

'------------------------------------------------------------------------------
' VBA fallback - same rules as SalesStaging.pq
'------------------------------------------------------------------------------
Public Sub TransformWithVBA()
    Dim loRaw As ListObject, loStage As ListObject
    Dim raw As Variant, nRaw As Long, rawMap As Collection
    Dim aliasMap As Collection, provMap As Collection, prodMap As Collection, seen As Collection
    Dim provInfo() As Variant, prodInfo() As Variant
    Dim rawCols() As Long, req As Variant, f() As String
    Dim keys() As Double, ties() As String, idx() As Long
    Dim out() As Variant, cleaned As Variant, provKey As String, dupKey As String
    Dim i As Long, j As Long, r As Long, n As Long, cId As Long, cSrc As Long

    Set loRaw = RequireTable("tblRaw")
    nRaw = TableRowCount(loRaw)
    Set loStage = EnsureStaticStagingTable()
    If nRaw = 0 Then
        WriteTable loStage, Empty, 0
        Exit Sub
    End If

    raw = TableToArray(loRaw)
    Set rawMap = HeaderMap(loRaw)
    req = ExportHeaders()
    ReDim rawCols(0 To UBound(req))
    For j = 0 To UBound(req)
        rawCols(j) = ColIndex(rawMap, CStr(req(j)), "tblRaw")
    Next j

    cId = ColIndex(rawMap, "Raw Row ID", "tblRaw")
    cSrc = ColIndex(rawMap, "Source File", "tblRaw")
    LoadLookups aliasMap, provMap, provInfo, prodMap, prodInfo

    ' process in Raw Row ID order so "keep the first duplicate" is predictable
    ReDim keys(1 To nRaw): ReDim ties(1 To nRaw): ReDim idx(1 To nRaw)
    For i = 1 To nRaw
        keys(i) = Val(ValueToText(raw(i, cId)))
        ties(i) = Format$(i, "000000000")
        idx(i) = i
    Next i
    QuickSortIndex keys, ties, idx, 1, nRaw

    Set seen = New Collection
    ReDim out(1 To nRaw, 1 To 24)
    ReDim f(0 To UBound(req))
    For r = 1 To nRaw
        i = idx(r)
        For j = 0 To UBound(req)
            f(j) = ValueToText(raw(i, rawCols(j)))
        Next j
        If Len(f(0)) > 0 Or Len(f(8)) > 0 Then                       ' skip blank lines
            cleaned = CleanRawRow(keys(i), ValueToText(raw(i, cSrc)), f, _
                                  aliasMap, provMap, provInfo, prodMap, prodInfo, provKey)
            dupKey = DuplicateKey(cleaned, provKey)
            If Not KeyExists(seen, dupKey) Then
                seen.Add 1, dupKey
                n = n + 1
                For j = 1 To 24
                    out(n, j) = cleaned(j)
                Next j
            End If
        End If
    Next r

    WriteTable loStage, out, n
    FormatColumn loStage, "OrderDate", "yyyy-mm-dd"
    FormatColumn loStage, "ShipDate", "yyyy-mm-dd"
End Sub

' Lookup tables -> collections keyed the same way the M code joins them.
' The collections hold row numbers into provInfo / prodInfo.
Public Sub LoadLookups(ByRef aliasMap As Collection, ByRef provMap As Collection, ByRef provInfo() As Variant, _
                       ByRef prodMap As Collection, ByRef prodInfo() As Variant)
    Dim lo As ListObject, data As Variant, map As Collection, i As Long, n As Long, key As String

    Set aliasMap = New Collection
    Set lo = RequireTable("tblProvinceAlias")
    data = TableToArray(lo)
    Set map = HeaderMap(lo)
    If Not IsEmpty(data) Then
        For i = 1 To UBound(data, 1)
            key = NormalizeKey(ValueToText(data(i, ColIndex(map, "Alias"))))
            If Not KeyExists(aliasMap, key) Then aliasMap.Add UCase$(ValueToText(data(i, ColIndex(map, "ProvinceCode")))), key
        Next i
    End If

    Set provMap = New Collection
    Set lo = RequireTable("tblProvinces")
    data = TableToArray(lo)
    Set map = HeaderMap(lo)
    n = 0
    If IsEmpty(data) Then
        ReDim provInfo(1 To 1, 1 To 2)
    Else
        ReDim provInfo(1 To UBound(data, 1), 1 To 2)
        For i = 1 To UBound(data, 1)
            key = UCase$(ValueToText(data(i, ColIndex(map, "ProvinceCode"))))
            If Not KeyExists(provMap, key) Then
                n = n + 1
                provInfo(n, 1) = data(i, ColIndex(map, "Province"))
                provInfo(n, 2) = data(i, ColIndex(map, "Region"))
                provMap.Add n, key
            End If
        Next i
    End If

    Set prodMap = New Collection
    Set lo = RequireTable("tblProducts")
    data = TableToArray(lo)
    Set map = HeaderMap(lo)
    n = 0
    If IsEmpty(data) Then
        ReDim prodInfo(1 To 1, 1 To 5)
    Else
        ReDim prodInfo(1 To UBound(data, 1), 1 To 5)
        For i = 1 To UBound(data, 1)
            key = UCase$(ValueToText(data(i, ColIndex(map, "ProductID"))))
            If Not KeyExists(prodMap, key) Then
                n = n + 1
                prodInfo(n, 1) = data(i, ColIndex(map, "ProductName"))
                prodInfo(n, 2) = data(i, ColIndex(map, "Category"))
                prodInfo(n, 3) = data(i, ColIndex(map, "SubCategory"))
                prodInfo(n, 4) = data(i, ColIndex(map, "ListPrice"))
                prodInfo(n, 5) = data(i, ColIndex(map, "UnitCost"))
                prodMap.Add n, key
            End If
        Next i
    End If
End Sub

Public Function StagingHeaders() As Variant
    StagingHeaders = Array("RawRowID", "SourceFile", "OrderID", "OrderDate", "ShipDate", "CustomerID", _
        "CustomerName", "Segment", "City", "ProvinceCode", "Province", "Region", "ProductID", "ProductName", _
        "Category", "SubCategory", "Quantity", "UnitPrice", "ListPrice", "UnitCost", "Discount", "Channel", _
        "SalesRep", "PaymentMethod")
End Function

' The fallback writes into a normal table - if tblStaging is the query table, swap it out
Private Function EnsureStaticStagingTable() As ListObject
    Dim ws As Worksheet, lo As ListObject, hdr As Variant

    Set lo = GetTable(STAGING_TABLE)
    If Not lo Is Nothing Then
        If lo.SourceType <> xlSrcQuery Then
            Set EnsureStaticStagingTable = lo
            Exit Function
        End If
        lo.Delete
    End If

    Set ws = ThisWorkbook.Worksheets(SH_STAGING)
    hdr = StagingHeaders()
    ws.Range("A3").Resize(1, UBound(hdr) + 1).Value = hdr
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A3").Resize(2, UBound(hdr) + 1), , xlYes)
    lo.Name = STAGING_TABLE
    lo.TableStyle = "TableStyleLight9"
    Set EnsureStaticStagingTable = lo
End Function
