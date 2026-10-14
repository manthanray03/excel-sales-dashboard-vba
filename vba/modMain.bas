Attribute VB_Name = "modMain"
'===============================================================================
' modMain
' Button entry points + the one-click refresh that runs the whole pipeline:
'
'   1. Import    new CSV exports in \data\raw          -> tblRaw         (modImport)
'   2. Clean     Power Query query "SalesStaging"      -> tblStaging     (modTransform)
'   3. Validate  rules from the Validation_Rules sheet -> tblExceptions  (modValidation)
'   4. Load      valid rows + calculated columns       -> tblSales       (modLoad)
'   5. Report    RFM, ABC, commission, pivots, slicers                   (modAnalytics, modDashboard)
'
' Every run is written to tblRunLog (Logs sheet) with its timings and counts.
'===============================================================================
Option Explicit

Public Const APP_TITLE As String = "Sales Performance Dashboard"

' sheet names - kept here so a rename is a one line change
Public Const SH_DASHBOARD As String = "Dashboard"
Public Const SH_GUIDE As String = "Guide"
Public Const SH_ENTRY As String = "Data_Entry"
Public Const SH_SALES As String = "Sales_Data"
Public Const SH_EXCEPTIONS As String = "Exceptions"
Public Const SH_RFM As String = "Customer_RFM"
Public Const SH_ABC As String = "Product_ABC"
Public Const SH_COMMISSION As String = "Commission"
Public Const SH_STAGING As String = "Staging"
Public Const SH_RAW As String = "Raw_Data"
Public Const SH_RULES As String = "Validation_Rules"
Public Const SH_LOOKUPS As String = "Lookups"
Public Const SH_LOGS As String = "Logs"
Public Const SH_PIVOTS As String = "Pivots"
Public Const SH_SETTINGS As String = "Settings"

'------------------------------------------------------------------------------
' Run Full Refresh button
'------------------------------------------------------------------------------
Public Sub RunFullRefresh()
    Dim started As Date, t0 As Double, stepName As String
    Dim imported As Long, method As String
    Dim res As ValidationResult
    Dim rowStatus() As Long

    started = Now
    t0 = Timer
    On Error GoTo Failed

    stepName = "import"
    SpeedOn "Step 1 of 5: importing new files..."
    imported = ImportNewFiles(False)

    stepName = "clean"
    SetStatus "Step 2 of 5: cleaning data (Power Query)..."
    If Not RunTransform(method) Then
        SpeedOff
        LogRun started, t0, imported, res, method, "Cancelled at clean step"
        Exit Sub
    End If

    stepName = "validate"
    SetStatus "Step 3 of 5: validating rows..."
    res = RunValidation(rowStatus)

    stepName = "load"
    SetStatus "Step 4 of 5: loading the sales table..."
    LoadSalesTable rowStatus

    stepName = "report"
    SetStatus "Step 5 of 5: analytics and dashboard..."
    BuildCustomerRFM
    BuildProductABC
    BuildCommission
    RefreshDashboard

    SaveRunStats imported, res, method
    LogRun started, t0, imported, res, method, "OK"
    SpeedOff
    ThisWorkbook.Worksheets(SH_DASHBOARD).Activate

    MsgBox "Refresh finished in " & Format$(ElapsedSeconds(t0), "0.0") & " seconds." & vbLf & vbLf & _
           "New rows imported:" & vbTab & Format$(imported, "#,##0") & vbLf & _
           "Rows validated:" & vbTab & Format$(res.TotalRows, "#,##0") & vbLf & _
           "Rows loaded:" & vbTab & Format$(res.PassedRows, "#,##0") & vbLf & _
           "Rows rejected:" & vbTab & Format$(res.FailedRows, "#,##0") & vbLf & _
           "Data accuracy:" & vbTab & Format$(AccuracyRate(res), "0.0%") & vbLf & vbLf & _
           "Cleaning step: " & method, vbInformation, APP_TITLE
    Exit Sub

Failed:
    Dim msg As String
    msg = Err.Description
    SpeedOff
    LogRun started, t0, imported, res, method, "Failed (" & stepName & "): " & msg
    MsgBox "The refresh stopped during the " & stepName & " step:" & vbLf & vbLf & msg, vbCritical, APP_TITLE
End Sub

Public Function AccuracyRate(ByRef res As ValidationResult) As Double
    If res.TotalRows > 0 Then AccuracyRate = res.PassedRows / res.TotalRows
End Function

'------------------------------------------------------------------------------
' Other buttons
'------------------------------------------------------------------------------
Public Sub GoToExceptions()
    ThisWorkbook.Worksheets(SH_EXCEPTIONS).Activate
    ThisWorkbook.Worksheets(SH_EXCEPTIONS).Range("A1").Select
End Sub

Public Sub GoToDashboard()
    ThisWorkbook.Worksheets(SH_DASHBOARD).Activate
    ThisWorkbook.Worksheets(SH_DASHBOARD).Range("A1").Select
End Sub

Public Sub GoToDataEntry()
    ThisWorkbook.Worksheets(SH_ENTRY).Activate
    ThisWorkbook.Worksheets(SH_ENTRY).Range("in_OrderID").Select
End Sub

' Wipes tblRaw + the import log and loads every file in the data folder again
Public Sub ResetAndReimportAll()
    If MsgBox("This clears Raw_Data and the import log, then re-imports every CSV in the data folder " & _
              "and runs a full refresh." & vbLf & vbLf & "Manual entries added through the Data Entry " & _
              "form will be removed too. Continue?", vbQuestion + vbYesNo + vbDefaultButton2, APP_TITLE) <> vbYes Then
        Exit Sub
    End If
    On Error GoTo Failed
    SpeedOn "Clearing raw data..."
    WriteTable RequireTable("tblRaw"), Empty, 0
    WriteTable RequireTable("tblImportLog"), Empty, 0
    SpeedOff
    RunFullRefresh
    Exit Sub
Failed:
    SpeedOff
    MsgBox "Reset failed: " & Err.Description, vbCritical, APP_TITLE
End Sub

'------------------------------------------------------------------------------
' Run stats + log
'------------------------------------------------------------------------------
Private Sub SaveRunStats(ByVal imported As Long, ByRef res As ValidationResult, ByVal method As String)
    Dim rawRows As Long
    rawRows = TableRowCount(RequireTable("tblRaw"))
    SetNamedValue "Stat_RawRows", rawRows
    SetNamedValue "Stat_Duplicates", rawRows - res.TotalRows
    SetNamedValue "Stat_Validated", res.TotalRows
    SetNamedValue "Stat_Loaded", res.PassedRows
    SetNamedValue "Stat_Rejected", res.FailedRows
    SetNamedValue "Stat_Warnings", res.WarningCount
    SetNamedValue "Stat_Accuracy", AccuracyRate(res)
    SetNamedValue "Stat_CleanMethod", method
    SetNamedValue "LastRefresh", Now
End Sub

Private Sub LogRun(ByVal started As Date, ByVal t0 As Double, ByVal imported As Long, _
                   ByRef res As ValidationResult, ByVal method As String, ByVal status As String)
    Dim lo As ListObject, logRow(1 To 1, 1 To 12) As Variant, lastId As Variant
    On Error Resume Next            ' logging must never be the thing that breaks a run
    Set lo = GetTable("tblRunLog")
    If lo Is Nothing Then Exit Sub
    lastId = 0
    If TableRowCount(lo) > 0 Then lastId = Application.WorksheetFunction.Max(lo.ListColumns("RunID").DataBodyRange)
    logRow(1, 1) = lastId + 1
    logRow(1, 2) = started
    logRow(1, 3) = Now
    logRow(1, 4) = ElapsedSeconds(t0)
    logRow(1, 5) = imported
    logRow(1, 6) = res.TotalRows
    logRow(1, 7) = res.PassedRows
    logRow(1, 8) = res.FailedRows
    logRow(1, 9) = res.WarningCount
    logRow(1, 10) = AccuracyRate(res)
    logRow(1, 11) = method
    logRow(1, 12) = status
    AppendToTable lo, logRow, 1
End Sub
