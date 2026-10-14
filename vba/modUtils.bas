Attribute VB_Name = "modUtils"
'===============================================================================
' modUtils
' Workbook helpers: speed mode, table read/write, named cells, paths.
'===============================================================================
Option Explicit

Private mSavedCalc As Long

'------------------------------------------------------------------------------
' Speed mode - switch off screen updates / events / recalcs during big writes
'------------------------------------------------------------------------------
Public Sub SpeedOn(Optional ByVal statusText As String = "")
    mSavedCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    If Len(statusText) > 0 Then Application.StatusBar = statusText
End Sub

Public Sub SpeedOff()
    Application.EnableEvents = True
    If mSavedCalc = 0 Then mSavedCalc = xlCalculationAutomatic
    Application.Calculation = mSavedCalc
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

Public Sub SetStatus(ByVal msg As String)
    Application.StatusBar = msg
    DoEvents
End Sub

'------------------------------------------------------------------------------
' Tables
'------------------------------------------------------------------------------
Public Function GetTable(ByVal tableName As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, tableName, vbTextCompare) = 0 Then
                Set GetTable = lo
                Exit Function
            End If
        Next lo
    Next ws
End Function

Public Function RequireTable(ByVal tableName As String) As ListObject
    Set RequireTable = GetTable(tableName)
    If RequireTable Is Nothing Then
        Err.Raise vbObjectError + 513, "RequireTable", "The table '" & tableName & "' is missing from the workbook."
    End If
End Function

' Number of data rows (a table with a single empty row counts as 0)
Public Function TableRowCount(ByVal lo As ListObject) As Long
    If lo.DataBodyRange Is Nothing Then Exit Function
    TableRowCount = lo.DataBodyRange.Rows.Count
    If TableRowCount = 1 Then
        If Application.WorksheetFunction.CountA(lo.DataBodyRange) = 0 Then TableRowCount = 0
    End If
End Function

' Table body as a 1-based 2D array (rows, columns). Empty when there are no rows.
Public Function TableToArray(ByVal lo As ListObject) As Variant
    Dim v As Variant, tmp() As Variant
    If TableRowCount(lo) = 0 Then
        TableToArray = Empty
        Exit Function
    End If
    v = lo.DataBodyRange.Value
    If Not IsArray(v) Then              ' 1 x 1 table
        ReDim tmp(1 To 1, 1 To 1)
        tmp(1, 1) = v
        v = tmp
    End If
    TableToArray = v
End Function

' Column name -> column number (case-insensitive) for one table
Public Function HeaderMap(ByVal lo As ListObject) As Collection
    Dim headers() As String, c As Long
    ReDim headers(1 To lo.ListColumns.Count)
    For c = 1 To lo.ListColumns.Count
        headers(c) = lo.ListColumns(c).Name
    Next c
    Set HeaderMap = MapFromHeaders(headers)
End Function

Public Function ColIndex(ByVal map As Collection, ByVal columnName As String, Optional ByVal tableName As String = "") As Long
    If KeyExists(map, LCase$(columnName)) Then
        ColIndex = map.Item(LCase$(columnName))
    Else
        Err.Raise vbObjectError + 514, "ColIndex", "Column '" & columnName & "' not found" & _
            IIf(Len(tableName) > 0, " in " & tableName, "") & "."
    End If
End Function

' Replaces everything in a table with data (1-based array, rowCount rows used).
' Number formats of the first data row are re-applied to the new rows.
' Values are written straight to the cells under the header and the table is
' then resized over them - an emptied table's DataBodyRange can be Nothing,
' even right after a Resize, so it isn't relied on here.
Public Sub WriteTable(ByVal lo As ListObject, ByRef data As Variant, ByVal rowCount As Long)
    Dim oldRows As Long, c As Long, nCols As Long, fmts() As String, target As Range

    nCols = lo.ListColumns.Count
    ReDim fmts(1 To nCols)
    For c = 1 To nCols
        fmts(c) = lo.HeaderRowRange.Cells(1, c).Offset(1, 0).NumberFormat
    Next c

    oldRows = TableRowCount(lo)
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.ClearContents

    If rowCount < 1 Then
        If lo.ListRows.Count > 1 Then
            c = lo.ListRows.Count
            lo.Resize lo.HeaderRowRange.Resize(2)
            lo.HeaderRowRange.Offset(2).Resize(c - 1).Clear
        End If
        Exit Sub
    End If

    EnsureDataRow lo
    Set target = lo.HeaderRowRange.Offset(1).Resize(rowCount, nCols)
    For c = 1 To nCols
        target.Columns(c).NumberFormat = fmts(c)
    Next c
    target.Value = data
    lo.Resize lo.HeaderRowRange.Resize(rowCount + 1)
    CheckRowCount lo, rowCount

    ' table got shorter - clear what's left underneath
    If oldRows > rowCount Then
        lo.HeaderRowRange.Offset(rowCount + 1).Resize(oldRows - rowCount).Clear
    End If
End Sub

' Adds rows to the bottom of a table (same write-then-resize approach as WriteTable)
Public Sub AppendToTable(ByVal lo As ListObject, ByRef data As Variant, ByVal rowCount As Long)
    Dim oldRows As Long, c As Long, nCols As Long, fmts() As String, target As Range
    If rowCount < 1 Then Exit Sub

    nCols = lo.ListColumns.Count
    ReDim fmts(1 To nCols)
    For c = 1 To nCols
        fmts(c) = lo.HeaderRowRange.Cells(1, c).Offset(1, 0).NumberFormat
    Next c

    EnsureDataRow lo
    oldRows = TableRowCount(lo)
    Set target = lo.HeaderRowRange.Offset(oldRows + 1).Resize(rowCount, nCols)
    For c = 1 To nCols
        target.Columns(c).NumberFormat = fmts(c)      ' set "@" before writing so raw text stays text
    Next c
    target.Value = data
    lo.Resize lo.HeaderRowRange.Resize(oldRows + rowCount + 1)
    CheckRowCount lo, oldRows + rowCount
End Sub

' An empty table can be left with only its "insert row" (no real data row);
' give it a proper row first so Resize behaves.
Private Sub EnsureDataRow(ByVal lo As ListObject)
    If lo.DataBodyRange Is Nothing Then lo.ListRows.Add
End Sub

Private Sub CheckRowCount(ByVal lo As ListObject, ByVal expected As Long)
    Dim actual As Long
    If Not lo.DataBodyRange Is Nothing Then actual = lo.DataBodyRange.Rows.Count
    If actual <> expected Then
        Err.Raise vbObjectError + 515, "WriteTable", "Table " & lo.Name & " should have " & expected & _
            " rows after the update but has " & actual & "."
    End If
End Sub

Public Sub FormatColumn(ByVal lo As ListObject, ByVal columnName As String, ByVal numberFormat As String)
    If lo.DataBodyRange Is Nothing Then Exit Sub
    lo.ListColumns(columnName).DataBodyRange.NumberFormat = numberFormat
End Sub

' Grows the first dimension of a (rows, cols) array, keeping the data
Public Sub GrowRows(ByRef arr() As Variant, ByVal newRows As Long)
    Dim tmp() As Variant, r As Long, c As Long
    ReDim tmp(1 To newRows, 1 To UBound(arr, 2))
    For r = 1 To UBound(arr, 1)
        For c = 1 To UBound(arr, 2)
            tmp(r, c) = arr(r, c)
        Next c
    Next r
    arr = tmp
End Sub

'------------------------------------------------------------------------------
' Named cells (Settings sheet)
'------------------------------------------------------------------------------
Public Function NamedValue(ByVal nm As String) As Variant
    On Error GoTo Missing
    NamedValue = ThisWorkbook.Names(nm).RefersToRange.Value
    Exit Function
Missing:
    NamedValue = Empty
End Function

Public Sub SetNamedValue(ByVal nm As String, ByVal v As Variant)
    ThisWorkbook.Names(nm).RefersToRange.Value = v
End Sub

'------------------------------------------------------------------------------
' Paths / platform
'------------------------------------------------------------------------------
Public Function PathSep() As String
    PathSep = Application.PathSeparator
End Function

Public Function IsMac() As Boolean
    IsMac = (InStr(1, Application.OperatingSystem, "Mac", vbTextCompare) > 0)
End Function

Public Function FolderExists(ByVal folderPath As String) As Boolean
    On Error GoTo Nope
    If Len(folderPath) = 0 Then Exit Function
    FolderExists = (Len(Dir(folderPath, vbDirectory)) > 0)
    Exit Function
Nope:
    FolderExists = False
End Function

Public Function EnsureFolder(ByVal folderPath As String) As Boolean
    On Error GoTo Nope
    If Not FolderExists(folderPath) Then MkDir folderPath
    EnsureFolder = True
    Exit Function
Nope:
    EnsureFolder = False
End Function

' Folder the workbook lives in (blank if it was never saved or lives on a web path)
Public Function WorkbookFolder() As String
    Dim p As String
    p = ThisWorkbook.Path
    If LCase$(Left$(p, 4)) = "http" Then p = ""
    WorkbookFolder = p
End Function

' Seconds between two Timer readings (Timer resets at midnight)
Public Function ElapsedSeconds(ByVal startTimer As Double) As Double
    Dim t As Double
    t = Timer - startTimer
    If t < 0 Then t = t + 86400
    ElapsedSeconds = Round(t, 1)
End Function
