Attribute VB_Name = "modDashboard"
'===============================================================================
' modDashboard
' Pivot tables (Pivots sheet), slicers (Dashboard), filter reset and PDF export.
'
' All pivots share ONE pivot cache built on tblSales - that's what lets a single
' slicer filter every chart and KPI at the same time. The KPI cards and chart
' ranges on the Dashboard/Calc sheets read the pivots with GETPIVOTDATA.
'===============================================================================
Option Explicit

Private Const SALES_TABLE As String = "tblSales"

Public Sub RefreshDashboard()
    EnsurePivotTables
    RefreshPivotCaches
    EnsureSlicers
    Application.Calculate
End Sub

Public Sub RefreshPivotCaches()
    Dim pc As PivotCache
    For Each pc In ThisWorkbook.PivotCaches
        pc.Refresh
    Next pc
End Sub

'------------------------------------------------------------------------------
' Pivot tables. These already exist in the workbook - this only rebuilds any that
' have been deleted, so the layout below doubles as documentation of each pivot.
'   name, top-left cell, row field, column field, data fields (Field:Caption;...)
'------------------------------------------------------------------------------
Public Sub EnsurePivotTables()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_PIVOTS)
    EnsurePivot ws, "ptRegion", "A3", "Region", "", "Revenue:Total Revenue;Profit:Total Profit;Quantity:Units Sold;OrderFlag:Orders"
    EnsurePivot ws, "ptCategory", "G3", "Category", "", "Revenue:Total Revenue;Profit:Total Profit"
    EnsurePivot ws, "ptChannel", "K3", "Channel", "", "Revenue:Total Revenue;OrderFlag:Orders"
    EnsurePivot ws, "ptTrend", "A14", "MonthNum", "Year", "Revenue:Total Revenue"
    EnsurePivot ws, "ptRep", "G14", "SalesRep", "", "Revenue:Total Revenue;Profit:Total Profit"
    EnsurePivot ws, "ptProduct", "K14", "ProductName", "", "Revenue:Total Revenue"
End Sub

Private Sub EnsurePivot(ByVal ws As Worksheet, ByVal ptName As String, ByVal topLeft As String, _
                        ByVal rowField As String, ByVal colField As String, ByVal dataSpec As String)
    Dim pt As PivotTable, specs As Variant, parts As Variant, i As Long, df As PivotField

    If PivotExists(ws, ptName) Then Exit Sub

    Set pt = SharedCache(ws).CreatePivotTable(TableDestination:=ws.Range(topLeft), TableName:=ptName)
    With pt
        .ManualUpdate = True
        If Len(rowField) > 0 Then .PivotFields(rowField).Orientation = xlRowField
        If Len(colField) > 0 Then .PivotFields(colField).Orientation = xlColumnField
        specs = Split(dataSpec, ";")
        For i = 0 To UBound(specs)
            parts = Split(specs(i), ":")
            Set df = .AddDataField(.PivotFields(parts(0)), parts(1), xlSum)
            df.NumberFormat = "#,##0"
        Next i
        .ManualUpdate = False
    End With
End Sub

' every pivot must use the same cache, otherwise slicers can't connect to all of them
Private Function SharedCache(ByVal ws As Worksheet) As PivotCache
    If ws.PivotTables.Count > 0 Then
        Set SharedCache = ws.PivotTables(1).PivotCache
    Else
        Set SharedCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SALES_TABLE)
    End If
End Function

Private Function PivotExists(ByVal ws As Worksheet, ByVal ptName As String) As Boolean
    Dim pt As PivotTable
    For Each pt In ws.PivotTables
        If StrComp(pt.Name, ptName, vbTextCompare) = 0 Then
            PivotExists = True
            Exit Function
        End If
    Next pt
End Function

'------------------------------------------------------------------------------
' Slicers - one per field, connected to every pivot on the Pivots sheet.
' Each slicer is sized to fill a named block of cells on the Dashboard.
'------------------------------------------------------------------------------
Public Sub EnsureSlicers()
    Dim wsD As Worksheet, wsP As Worksheet
    Set wsD = ThisWorkbook.Worksheets(SH_DASHBOARD)
    Set wsP = ThisWorkbook.Worksheets(SH_PIVOTS)
    If wsP.PivotTables.Count = 0 Then Exit Sub

    AddSlicer wsD, wsP, "Year", "slicer_Year", 2
    AddSlicer wsD, wsP, "Region", "slicer_Region", 1
    AddSlicer wsD, wsP, "Category", "slicer_Category", 1
    AddSlicer wsD, wsP, "Channel", "slicer_Channel", 2
End Sub

Private Sub AddSlicer(ByVal wsD As Worksheet, ByVal wsP As Worksheet, ByVal fieldName As String, _
                      ByVal areaName As String, ByVal nCols As Long)
    Dim sc As SlicerCache, sl As Slicer, area As Range, i As Long

    If Not FindSlicerCache(fieldName) Is Nothing Then Exit Sub
    Set area = wsD.Range(areaName)

    Set sc = NewSlicerCache(wsP.PivotTables(1), fieldName)
    For i = 2 To wsP.PivotTables.Count
        sc.PivotTables.AddPivotTable wsP.PivotTables(i)
    Next i

    Set sl = sc.Slicers.Add(wsD, , "slc" & fieldName, fieldName, area.Top, area.Left, area.Width, area.Height)
    sl.Style = "SlicerStyleLight1"
    sl.NumberOfColumns = nCols
End Sub

' SlicerCaches.Add2 is Excel 2013+, Add is the older (2010) name
Private Function NewSlicerCache(ByVal pt As PivotTable, ByVal fieldName As String) As SlicerCache
    On Error Resume Next
    Set NewSlicerCache = ThisWorkbook.SlicerCaches.Add2(pt, fieldName)
    If NewSlicerCache Is Nothing Then
        Err.Clear
        Set NewSlicerCache = ThisWorkbook.SlicerCaches.Add(pt, fieldName)
    End If
    On Error GoTo 0
    If NewSlicerCache Is Nothing Then Err.Raise vbObjectError + 540, , "Couldn't create the " & fieldName & " slicer."
End Function

Private Function FindSlicerCache(ByVal fieldName As String) As SlicerCache
    Dim sc As SlicerCache
    For Each sc In ThisWorkbook.SlicerCaches
        If StrComp(sc.SourceName, fieldName, vbTextCompare) = 0 Then
            Set FindSlicerCache = sc
            Exit Function
        End If
    Next sc
End Function

'------------------------------------------------------------------------------
' Clear Filters button
'------------------------------------------------------------------------------
Public Sub ClearAllFilters()
    Dim sc As SlicerCache
    On Error GoTo Failed
    For Each sc In ThisWorkbook.SlicerCaches
        sc.ClearManualFilter
    Next sc
    Exit Sub
Failed:
    MsgBox "Couldn't clear the filters: " & Err.Description, vbExclamation, APP_TITLE
End Sub

'------------------------------------------------------------------------------
' Export PDF button - saves the dashboard (as currently filtered) to \reports
'------------------------------------------------------------------------------
Public Sub ExportDashboardPDF()
    Dim folder As String, filePath As String
    On Error GoTo Failed

    folder = WorkbookFolder()
    If Len(folder) = 0 Then
        MsgBox "Save the workbook to a folder on this computer first.", vbInformation, APP_TITLE
        Exit Sub
    End If
    If EnsureFolder(folder & PathSep() & "reports") Then folder = folder & PathSep() & "reports"
    filePath = folder & PathSep() & "Sales_Dashboard_" & Format$(Now, "yyyy-mm-dd_hhnn") & ".pdf"

    ThisWorkbook.Worksheets(SH_DASHBOARD).ExportAsFixedFormat Type:=xlTypePDF, Filename:=filePath, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    MsgBox "Dashboard saved to:" & vbLf & filePath, vbInformation, APP_TITLE
    Exit Sub
Failed:
    MsgBox "Couldn't export the PDF: " & Err.Description, vbExclamation, APP_TITLE
End Sub
