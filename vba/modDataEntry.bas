Attribute VB_Name = "modDataEntry"
'===============================================================================
' modDataEntry
' The Data_Entry sheet is a standard form for orders that never made it into a
' POS export (phone orders, corrections...). Dropdowns + Excel data validation
' catch most mistakes while typing; SubmitEntry double-checks everything and
' appends the order to tblRaw in the same format as an export, so it goes
' through exactly the same cleaning and validation as everything else.
'
' CreateUploadTemplate builds the blank Excel template stores use for their
' monthly export when the POS can't produce one.
'===============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Submit button
'------------------------------------------------------------------------------
Public Sub SubmitEntry()
    Dim problems As String, lo As ListObject, rawRow(1 To 1, 1 To 18) As Variant
    Dim od As Variant, sd As Variant, qty As Variant, price As Variant, disc As Variant
    Dim orderId As String, productId As String, nextId As Long

    orderId = UCase$(InputText("in_OrderID"))
    productId = UCase$(InputText("in_ProductID"))
    od = ToDateValue(NamedValue("in_OrderDate"))
    sd = ToDateValue(NamedValue("in_ShipDate"))
    qty = NamedValue("in_Quantity")
    price = NamedValue("in_UnitPrice")
    disc = NamedValue("in_Discount")
    If IsBlankValue(disc) Then disc = 0

    If Len(orderId) = 0 Then problems = problems & vbLf & "- Order ID is blank"
    If IsEmpty(od) Then problems = problems & vbLf & "- Order date is missing or not a date"
    If IsEmpty(sd) Then
        problems = problems & vbLf & "- Ship date is missing or not a date"
    ElseIf Not IsEmpty(od) Then
        If sd < od Then problems = problems & vbLf & "- Ship date is before the order date"
    End If
    If Len(InputText("in_CustomerID")) = 0 Then problems = problems & vbLf & "- Customer ID is blank"
    If Len(InputText("in_CustomerName")) = 0 Then problems = problems & vbLf & "- Customer name is blank"
    If Len(InputText("in_Segment")) = 0 Then problems = problems & vbLf & "- Pick a segment"
    If Len(InputText("in_City")) = 0 Then problems = problems & vbLf & "- City is blank"
    If Len(InputText("in_Province")) = 0 Then problems = problems & vbLf & "- Pick a province"
    If Len(productId) = 0 Then problems = problems & vbLf & "- Pick a product"
    If Not IsNumberValue(qty) Then
        problems = problems & vbLf & "- Quantity must be a number"
    ElseIf qty < 1 Or qty > 100 Or qty <> Int(qty) Then
        problems = problems & vbLf & "- Quantity must be a whole number from 1 to 100"
    End If
    If Not IsNumberValue(price) Then
        problems = problems & vbLf & "- Unit price must be a number"
    ElseIf price <= 0 Then
        problems = problems & vbLf & "- Unit price must be more than 0"
    End If
    If Not IsNumberValue(disc) Then
        problems = problems & vbLf & "- Discount must be a number (e.g. 10%)"
    ElseIf disc < 0 Or disc > 0.5 Then
        problems = problems & vbLf & "- Discount must be between 0% and 50%"
    End If
    If Len(InputText("in_Channel")) = 0 Then problems = problems & vbLf & "- Pick a sales channel"
    If Len(InputText("in_SalesRep")) = 0 Then problems = problems & vbLf & "- Pick a sales rep"
    If Len(InputText("in_Payment")) = 0 Then problems = problems & vbLf & "- Pick a payment method"

    If Len(problems) > 0 Then
        ShowStatus "Not submitted - fix the highlighted fields.", False
        MsgBox "Please fix these before submitting:" & vbLf & problems, vbExclamation, APP_TITLE
        Exit Sub
    End If

    Set lo = RequireTable("tblRaw")
    If LineAlreadyEntered(lo, orderId, productId) Then
        If MsgBox(orderId & " already has a line for " & productId & "." & vbLf & _
                  "Add it anyway? (validation will flag it as a duplicate)", vbQuestion + vbYesNo, APP_TITLE) <> vbYes Then
            Exit Sub
        End If
    End If

    If TableRowCount(lo) = 0 Then
        nextId = 1
    Else
        nextId = Application.WorksheetFunction.Max(lo.ListColumns("Raw Row ID").DataBodyRange) + 1
    End If

    ' same layout and formats as a POS export line
    rawRow(1, 1) = nextId
    rawRow(1, 2) = "Manual Entry"
    rawRow(1, 3) = orderId
    rawRow(1, 4) = Format$(od, "yyyy-mm-dd")
    rawRow(1, 5) = Format$(sd, "yyyy-mm-dd")
    rawRow(1, 6) = UCase$(InputText("in_CustomerID"))
    rawRow(1, 7) = InputText("in_CustomerName")
    rawRow(1, 8) = InputText("in_Segment")
    rawRow(1, 9) = InputText("in_City")
    rawRow(1, 10) = InputText("in_Province")
    rawRow(1, 11) = productId
    rawRow(1, 12) = InputText("in_Product")
    rawRow(1, 13) = InvariantNumber(qty)
    rawRow(1, 14) = InvariantNumber(price)
    rawRow(1, 15) = InvariantNumber(disc)
    rawRow(1, 16) = InputText("in_Channel")
    rawRow(1, 17) = InputText("in_SalesRep")
    rawRow(1, 18) = InputText("in_Payment")

    On Error GoTo Failed
    Application.EnableEvents = False
    AppendToTable lo, rawRow, 1
    Application.EnableEvents = True

    ClearEntryForm True
    ShowStatus "Added " & orderId & " / " & productId & " to Raw_Data (row " & nextId & "). " & _
               "Run Full Refresh to see it on the dashboard.", True
    Exit Sub
Failed:
    Application.EnableEvents = True
    MsgBox "Couldn't add the entry: " & Err.Description, vbCritical, APP_TITLE
End Sub

' Clear button (keepDates keeps the order/ship date - handy when keying a batch)
Public Sub ClearEntryForm(Optional ByVal keepDates As Boolean = False)
    Dim fieldNames As Variant, i As Long
    fieldNames = Array("in_OrderID", "in_CustomerID", "in_CustomerName", "in_Segment", "in_City", _
                       "in_Province", "in_Product", "in_Quantity", "in_UnitPrice", "in_Discount", _
                       "in_Channel", "in_SalesRep", "in_Payment")
    Application.EnableEvents = False
    For i = 0 To UBound(fieldNames)
        ThisWorkbook.Names(fieldNames(i)).RefersToRange.ClearContents
    Next i
    If Not keepDates Then
        ThisWorkbook.Names("in_OrderDate").RefersToRange.ClearContents
        ThisWorkbook.Names("in_ShipDate").RefersToRange.ClearContents
        ShowStatus "", True
    End If
    Application.EnableEvents = True
End Sub

Public Sub ClearEntryFormButton()
    ClearEntryForm False
End Sub

' Called from the Data_Entry sheet when a product is picked - fills in the list price
Public Sub FillListPrice()
    Dim lo As ListObject, productNames As Variant, prices As Variant, picked As String, i As Long
    picked = InputText("in_Product")
    If Len(picked) = 0 Then Exit Sub
    Set lo = RequireTable("tblProducts")
    If TableRowCount(lo) = 0 Then Exit Sub
    productNames = lo.ListColumns("ProductName").DataBodyRange.Value
    prices = lo.ListColumns("ListPrice").DataBodyRange.Value
    For i = 1 To UBound(productNames, 1)
        If StrComp(CStr(productNames(i, 1)), picked, vbTextCompare) = 0 Then
            Application.EnableEvents = False
            ThisWorkbook.Names("in_UnitPrice").RefersToRange.Value = prices(i, 1)
            Application.EnableEvents = True
            Exit Sub
        End If
    Next i
End Sub

'------------------------------------------------------------------------------
' Blank upload template for stores (saved to \templates)
'------------------------------------------------------------------------------
Public Sub CreateUploadTemplate()
    Dim wb As Workbook, ws As Worksheet, wsList As Worksheet, wsHelp As Worksheet
    Dim hdr As Variant, folder As String, filePath As String

    folder = WorkbookFolder()
    If Len(folder) = 0 Then
        MsgBox "Save this workbook to a folder first.", vbInformation, APP_TITLE
        Exit Sub
    End If
    If EnsureFolder(folder & PathSep() & "templates") Then folder = folder & PathSep() & "templates"
    filePath = folder & PathSep() & "Sales_Upload_Template.xlsx"

    On Error GoTo Failed
    Application.ScreenUpdating = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Name = "Sales Upload"
    Set wsList = wb.Worksheets.Add(After:=ws)
    wsList.Name = "Lists"
    Set wsHelp = wb.Worksheets.Add(After:=wsList)
    wsHelp.Name = "How to fill"

    ' header row, same columns the importer expects
    hdr = ExportHeaders()
    With ws.Range("A1").Resize(1, UBound(hdr) + 1)
        .Value = hdr
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 58, 95)
    End With
    ws.Range("A2:P1001").NumberFormat = "@"           ' keep everything as typed
    ws.Columns("A:P").ColumnWidth = 16
    ws.Activate
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True

    ' dropdown lists copied from the lookups
    CopyColumnTo "tblProvinces", "ProvinceCode", wsList, 1
    CopyColumnTo "tblProducts", "ProductID", wsList, 2
    CopyColumnTo "tblSegments", "Segment", wsList, 3
    CopyColumnTo "tblChannels", "Channel", wsList, 4
    CopyColumnTo "tblPayments", "PaymentMethod", wsList, 5
    CopyColumnTo "tblReps", "SalesRep", wsList, 6

    AddListValidation ws.Range("F2:F1001"), wsList, 3   ' Segment
    AddListValidation ws.Range("H2:H1001"), wsList, 1   ' Province
    AddListValidation ws.Range("I2:I1001"), wsList, 2   ' Product ID
    AddListValidation ws.Range("N2:N1001"), wsList, 4   ' Sales Channel
    AddListValidation ws.Range("O2:O1001"), wsList, 6   ' Sales Rep
    AddListValidation ws.Range("P2:P1001"), wsList, 5   ' Payment Method

    With wsHelp
        .Range("A1").Value = "How to fill in the sales upload template"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A3").Value = "1. One row per product line. Lines from the same order share the Order ID."
        .Range("A4").Value = "2. Dates as YYYY-MM-DD (e.g. 2025-03-07)."
        .Range("A5").Value = "3. Quantity is a whole number. Unit Price without the $ sign. Discount as 10% or 0.10 (blank = none)."
        .Range("A6").Value = "4. Use the dropdowns for Segment, Province, Product ID, Channel, Sales Rep and Payment Method."
        .Range("A7").Value = "5. Save as CSV (Comma delimited) named sales_YYYY_MM.csv and drop it in the data\raw folder."
        .Columns("A").ColumnWidth = 110
    End With

    Application.DisplayAlerts = False
    wb.SaveAs Filename:=filePath, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    wb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    MsgBox "Template saved to:" & vbLf & filePath, vbInformation, APP_TITLE
    Exit Sub

Failed:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Couldn't create the template: " & Err.Description, vbCritical, APP_TITLE
End Sub

Private Sub CopyColumnTo(ByVal tableName As String, ByVal columnName As String, ByVal wsList As Worksheet, ByVal targetCol As Long)
    Dim lo As ListObject, n As Long
    Set lo = RequireTable(tableName)
    n = TableRowCount(lo)
    wsList.Cells(1, targetCol).Value = columnName
    If n > 0 Then wsList.Cells(2, targetCol).Resize(n, 1).Value = lo.ListColumns(columnName).DataBodyRange.Value
End Sub

Private Sub AddListValidation(ByVal target As Range, ByVal wsList As Worksheet, ByVal listCol As Long)
    Dim lastRow As Long, src As String
    lastRow = wsList.Cells(wsList.Rows.Count, listCol).End(xlUp).Row
    src = "='" & wsList.Name & "'!" & wsList.Range(wsList.Cells(2, listCol), wsList.Cells(lastRow, listCol)).Address
    With target.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=src
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
End Sub

'------------------------------------------------------------------------------
' small helpers
'------------------------------------------------------------------------------
Private Function InputText(ByVal nm As String) As String
    InputText = CleanSpaces(ValueToText(NamedValue(nm)))
End Function

' number as text with a "." decimal point whatever the PC's settings are
Private Function InvariantNumber(ByVal v As Variant) As String
    Dim s As String
    s = Trim$(Str$(CDbl(v)))                ' Str$ always uses "." but drops the leading zero
    If Left$(s, 1) = "." Then s = "0" & s
    If Left$(s, 2) = "-." Then s = "-0" & Mid$(s, 2)
    InvariantNumber = s
End Function

Private Function LineAlreadyEntered(ByVal lo As ListObject, ByVal orderId As String, ByVal productId As String) As Boolean
    Dim data As Variant, map As Collection, i As Long, cO As Long, cP As Long
    data = TableToArray(lo)
    If IsEmpty(data) Then Exit Function
    Set map = HeaderMap(lo)
    cO = ColIndex(map, "Order ID")
    cP = ColIndex(map, "Product ID")
    For i = 1 To UBound(data, 1)
        If UCase$(Trim$(CStr(data(i, cO)))) = orderId Then
            If UCase$(Trim$(CStr(data(i, cP)))) = productId Then
                LineAlreadyEntered = True
                Exit Function
            End If
        End If
    Next i
End Function

Private Sub ShowStatus(ByVal msg As String, ByVal ok As Boolean)
    With ThisWorkbook.Names("entry_Status").RefersToRange
        .Value = msg
        If ok Then .Font.Color = RGB(30, 120, 70) Else .Font.Color = RGB(190, 40, 40)
    End With
End Sub
