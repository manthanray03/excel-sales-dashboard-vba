Attribute VB_Name = "modImport"
'===============================================================================
' modImport  -  "E" of the ETL
' Picks up the monthly POS exports (CSV) from the data folder and appends them
' to tblRaw exactly as they are (all text). Files already listed in the import
' log are skipped, so the button can be pressed any time.
'===============================================================================
Option Explicit

Private Const RAW_TABLE As String = "tblRaw"
Private Const LOG_TABLE As String = "tblImportLog"

' columns every export must have, in the order they are stored in tblRaw
Public Function ExportHeaders() As Variant
    ExportHeaders = Array("Order ID", "Order Date", "Ship Date", "Customer ID", "Customer Name", _
        "Segment", "City", "Province", "Product ID", "Product Name", "Quantity", "Unit Price", _
        "Discount", "Sales Channel", "Sales Rep", "Payment Method")
End Function

'------------------------------------------------------------------------------
' Import New Files button
'------------------------------------------------------------------------------
Public Sub ImportNewFilesButton()
    On Error GoTo Failed
    SpeedOn "Importing new files..."
    ImportNewFiles True
    SpeedOff
    Exit Sub
Failed:
    SpeedOff
    MsgBox "Import failed: " & Err.Description, vbCritical, APP_TITLE
End Sub

' Returns the number of rows added to tblRaw
Public Function ImportNewFiles(ByVal showMessage As Boolean) As Long
    Dim folder As String, files As Collection, f As Variant
    Dim done As Collection, loRaw As ListObject, loLog As ListObject
    Dim newRows As Variant, nRows As Long, nextId As Long, total As Long
    Dim summary As String, skipped As Long, errMsg As String, nFiles As Long

    Set loRaw = RequireTable(RAW_TABLE)
    Set loLog = RequireTable(LOG_TABLE)

    folder = GetDataFolder()
    If Len(folder) = 0 Then
        If showMessage Then MsgBox "No folder selected - nothing imported.", vbInformation, APP_TITLE
        Exit Function
    End If

    Set files = ListCsvFiles(folder)
    Set done = ImportedFileNames(loLog)
    nextId = NextRawRowId(loRaw)

    For Each f In files
        If KeyExists(done, LCase$(CStr(f))) Then
            skipped = skipped + 1
        Else
            errMsg = ""
            newRows = ParseExportFile(folder & PathSep() & f, CStr(f), nextId, nRows, errMsg)
            If Len(errMsg) > 0 Then
                AddImportLog loLog, CStr(f), 0, "Skipped - " & errMsg
                summary = summary & vbLf & "   " & f & "  -  skipped (" & errMsg & ")"
            Else
                If nRows > 0 Then AppendToTable loRaw, newRows, nRows
                AddImportLog loLog, CStr(f), nRows, "Imported"
                nextId = nextId + nRows
                total = total + nRows
                nFiles = nFiles + 1
                summary = summary & vbLf & "   " & f & "  -  " & Format$(nRows, "#,##0") & " rows"
            End If
        End If
    Next f

    If showMessage Then
        If nFiles = 0 And Len(summary) = 0 Then
            MsgBox "No new files found in" & vbLf & folder & vbLf & vbLf & _
                   skipped & " file(s) were already imported.", vbInformation, APP_TITLE
        Else
            MsgBox "Imported " & Format$(total, "#,##0") & " rows from " & nFiles & " file(s):" & summary & _
                   vbLf & vbLf & "Run Full Refresh to push them through to the dashboard.", vbInformation, APP_TITLE
        End If
    End If
    ImportNewFiles = total
End Function

'------------------------------------------------------------------------------
' Folder handling
'------------------------------------------------------------------------------

' Settings!DataFolder if filled in, otherwise <workbook folder>\data\raw.
' Asks the user to point at the folder if neither can be found.
Public Function GetDataFolder() As String
    Dim folder As String, picked As Variant

    folder = Trim$(CStr(NamedValue("DataFolder")))
    If Len(folder) = 0 And Len(WorkbookFolder()) > 0 Then
        folder = WorkbookFolder() & PathSep() & "data" & PathSep() & "raw"
    End If
    If Right$(folder, 1) = PathSep() Then folder = Left$(folder, Len(folder) - 1)

    If FolderExists(folder) Then
        GetDataFolder = folder
        Exit Function
    End If

    ' moved, opened from OneDrive/email, or a Mac sandbox thing - ask for a file in the folder
    MsgBox "Couldn't find the data folder:" & vbLf & IIf(Len(folder) = 0, "(workbook not saved yet)", folder) & _
           vbLf & vbLf & "Select any CSV file inside the folder that holds the monthly exports.", _
           vbInformation, APP_TITLE
    If IsMac() Then
        picked = Application.GetOpenFilename()
    Else
        picked = Application.GetOpenFilename("CSV files (*.csv),*.csv", , "Select a file in the data folder")
    End If
    If VarType(picked) = vbBoolean Then Exit Function      ' cancelled

    folder = Left$(picked, InStrRev(picked, PathSep()) - 1)
    SetNamedValue "DataFolder", folder                      ' remember it for next time
    GetDataFolder = folder
End Function

' All *.csv files in the folder, sorted by name (sales_2024_01, sales_2024_02, ...)
Private Function ListCsvFiles(ByVal folder As String) As Collection
    Dim fileNames() As String, n As Long, f As String, i As Long, j As Long, tmp As String
    Dim col As New Collection

    ReDim fileNames(1 To 64)
    f = Dir(folder & PathSep())         ' no wildcard - Dir("*.csv") doesn't work on a Mac
    Do While Len(f) > 0
        If LCase$(Right$(f, 4)) = ".csv" Then
            n = n + 1
            If n > UBound(fileNames) Then
                ReDim Preserve fileNames(1 To n * 2)
            End If
            fileNames(n) = f
        End If
        f = Dir()
    Loop

    ' Dir doesn't guarantee any order - insertion sort by name
    For i = 2 To n
        tmp = fileNames(i)
        j = i - 1
        Do While j >= 1
            If StrComp(fileNames(j), tmp, vbTextCompare) <= 0 Then Exit Do
            fileNames(j + 1) = fileNames(j)
            j = j - 1
        Loop
        fileNames(j + 1) = tmp
    Next i

    For i = 1 To n
        col.Add fileNames(i)
    Next i
    Set ListCsvFiles = col
End Function

'------------------------------------------------------------------------------
' Reading one export
'------------------------------------------------------------------------------
Private Function ReadTextFile(ByVal filePath As String) As String
    Dim fNum As Integer, s As String
    fNum = FreeFile
    Open filePath For Binary Access Read As #fNum
    s = Space$(LOF(fNum))
    If Len(s) > 0 Then Get #fNum, , s
    Close #fNum
    If Left$(s, 3) = Chr$(239) & Chr$(187) & Chr$(191) Then s = Mid$(s, 4)    ' UTF-8 BOM
    ReadTextFile = s
End Function

' Returns a (rows, 18) array: Raw Row ID, Source File, then the 16 export columns.
' Columns are matched by header name, so the column order in the file doesn't matter.
Private Function ParseExportFile(ByVal filePath As String, ByVal fileName As String, ByVal firstId As Long, _
                                 ByRef rowCount As Long, ByRef errMsg As String) As Variant
    Dim fileText As String, lines() As String, header As Variant, fields As Variant
    Dim req As Variant, colPos() As Long, out() As Variant
    Dim i As Long, j As Long, k As Long, n As Long, lineText As String

    rowCount = 0
    fileText = ReadTextFile(filePath)
    fileText = Replace(fileText, vbCrLf, vbLf)
    fileText = Replace(fileText, vbCr, vbLf)
    lines = Split(fileText, vbLf)
    If UBound(lines) < 0 Then
        errMsg = "file is empty"
        Exit Function
    End If

    ' map required columns to their position in this file
    header = SplitCsvLine(lines(0))
    req = ExportHeaders()
    ReDim colPos(0 To UBound(req))
    For j = 0 To UBound(req)
        colPos(j) = -1
        For k = 0 To UBound(header)
            If StrComp(Trim$(header(k)), req(j), vbTextCompare) = 0 Then
                colPos(j) = k
                Exit For
            End If
        Next k
        If colPos(j) = -1 Then
            errMsg = "missing column '" & req(j) & "'"
            Exit Function
        End If
    Next j

    If UBound(lines) < 1 Then
        errMsg = "no data rows"
        Exit Function
    End If

    ReDim out(1 To UBound(lines), 1 To UBound(req) + 3)
    i = 1
    Do While i <= UBound(lines)
        lineText = lines(i)
        ' a quoted field with a line break in it - glue the next line back on
        Do While (QuoteCount(lineText) Mod 2 = 1) And i < UBound(lines)
            i = i + 1
            lineText = lineText & vbLf & lines(i)
        Loop
        If Len(Trim$(lineText)) > 0 Then
            fields = SplitCsvLine(lineText)
            If Not AllFieldsBlank(fields) Then
                n = n + 1
                out(n, 1) = firstId + n - 1
                out(n, 2) = fileName
                For j = 0 To UBound(req)
                    If colPos(j) <= UBound(fields) Then
                        out(n, j + 3) = fields(colPos(j))
                    Else
                        out(n, j + 3) = ""
                    End If
                Next j
            End If
        End If
        i = i + 1
    Loop

    rowCount = n
    ParseExportFile = out
End Function

'------------------------------------------------------------------------------
' Import log
'------------------------------------------------------------------------------
Private Function ImportedFileNames(ByVal loLog As ListObject) As Collection
    Dim col As New Collection, data As Variant, i As Long, key As String
    data = TableToArray(loLog)
    If Not IsEmpty(data) Then
        For i = 1 To UBound(data, 1)
            If LCase$(CStr(data(i, 4))) = "imported" Then
                key = LCase$(CStr(data(i, 1)))
                If Not KeyExists(col, key) Then col.Add key, key
            End If
        Next i
    End If
    Set ImportedFileNames = col
End Function

Private Sub AddImportLog(ByVal loLog As ListObject, ByVal fileName As String, ByVal nRows As Long, ByVal status As String)
    Dim logRow(1 To 1, 1 To 4) As Variant
    logRow(1, 1) = fileName
    logRow(1, 2) = nRows
    logRow(1, 3) = Now
    logRow(1, 4) = status
    AppendToTable loLog, logRow, 1
End Sub

Private Function NextRawRowId(ByVal loRaw As ListObject) As Long
    If TableRowCount(loRaw) = 0 Then
        NextRawRowId = 1
    Else
        NextRawRowId = Application.WorksheetFunction.Max(loRaw.ListColumns("Raw Row ID").DataBodyRange) + 1
    End If
End Function
