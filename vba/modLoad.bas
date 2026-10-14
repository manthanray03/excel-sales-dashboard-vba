Attribute VB_Name = "modLoad"
'===============================================================================
' modLoad  -  "L" of the ETL
' Copies the rows that passed validation from tblStaging into tblSales (the table
' every pivot and chart reads from) and adds the calculated columns:
' date parts, ship days, gross sales, discount, revenue, COGS, profit, margin,
' and an OrderFlag (1 on the first line of each order) so pivots can count
' orders without the Data Model.
'===============================================================================
Option Explicit

Public Sub LoadSalesTable(ByRef rowStatus() As Long)
    Dim loStage As ListObject, data As Variant, out() As Variant, n As Long, nRows As Long

    Set loStage = RequireTable(STAGING_TABLE)
    nRows = TableRowCount(loStage)
    If nRows > 0 Then
        data = TableToArray(loStage)
        n = BuildSalesRows(data, HeaderMap(loStage), nRows, rowStatus, out)
    End If
    If n = 0 Then
        WriteTable RequireTable("tblSales"), Empty, 0
    Else
        WriteTable RequireTable("tblSales"), out, n
    End If
End Sub

' Rows that passed validation (rowStatus < 2) + calculated columns, in tblSales column order.
' Returns the number of rows put in out().
Public Function BuildSalesRows(ByRef data As Variant, ByVal map As Collection, ByVal nRows As Long, _
                               ByRef rowStatus() As Long, ByRef out() As Variant) As Long
    Dim nValid As Long, i As Long, n As Long
    Dim seenOrders As Collection, monthNames As Variant
    Dim od As Date, sd As Date, qty As Double, price As Double, disc As Double, cost As Double
    Dim gross As Double, discAmt As Double, revenue As Double, cogs As Double, profit As Double
    Dim orderId As String
    Dim cRaw As Long, cOrder As Long, cOD As Long, cSD As Long, cCust As Long, cName As Long
    Dim cSeg As Long, cCity As Long, cProv As Long, cRegion As Long, cProd As Long, cProdName As Long
    Dim cCat As Long, cSub As Long, cChan As Long, cRep As Long, cPay As Long
    Dim cQty As Long, cPrice As Long, cDisc As Long, cCost As Long

    For i = 1 To nRows
        If rowStatus(i) < 2 Then nValid = nValid + 1
    Next i
    If nValid = 0 Then Exit Function

    cRaw = FieldIndex(map, "RawRowID"):       cOrder = FieldIndex(map, "OrderID")
    cOD = FieldIndex(map, "OrderDate"):       cSD = FieldIndex(map, "ShipDate")
    cCust = FieldIndex(map, "CustomerID"):    cName = FieldIndex(map, "CustomerName")
    cSeg = FieldIndex(map, "Segment"):        cCity = FieldIndex(map, "City")
    cProv = FieldIndex(map, "Province"):      cRegion = FieldIndex(map, "Region")
    cProd = FieldIndex(map, "ProductID"):     cProdName = FieldIndex(map, "ProductName")
    cCat = FieldIndex(map, "Category"):       cSub = FieldIndex(map, "SubCategory")
    cChan = FieldIndex(map, "Channel"):       cRep = FieldIndex(map, "SalesRep")
    cPay = FieldIndex(map, "PaymentMethod"):  cQty = FieldIndex(map, "Quantity")
    cPrice = FieldIndex(map, "UnitPrice"):    cDisc = FieldIndex(map, "Discount")
    cCost = FieldIndex(map, "UnitCost")

    monthNames = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
    Set seenOrders = New Collection
    ReDim out(1 To nValid, 1 To 33)

    For i = 1 To nRows
        If rowStatus(i) < 2 Then
            n = n + 1
            od = data(i, cOD)
            sd = data(i, cSD)
            qty = data(i, cQty)
            price = data(i, cPrice)
            disc = data(i, cDisc)
            cost = data(i, cCost)

            gross = Round2(qty * price)
            discAmt = Round2(gross * disc)
            revenue = Round2(gross - discAmt)
            cogs = Round2(qty * cost)
            profit = Round2(revenue - cogs)

            orderId = CStr(data(i, cOrder))

            out(n, 1) = data(i, cRaw)                       ' LineID
            out(n, 2) = orderId
            out(n, 3) = od
            out(n, 4) = Year(od)
            out(n, 5) = "Q" & ((Month(od) - 1) \ 3 + 1)
            out(n, 6) = Month(od)
            out(n, 7) = monthNames(Month(od) - 1)            ' fixed English names, not locale dependent
            out(n, 8) = sd
            out(n, 9) = CLng(sd - od)                        ' ShipDays
            out(n, 10) = data(i, cCust)
            out(n, 11) = data(i, cName)
            out(n, 12) = data(i, cSeg)
            out(n, 13) = data(i, cCity)
            out(n, 14) = data(i, cProv)
            out(n, 15) = data(i, cRegion)
            out(n, 16) = data(i, cProd)
            out(n, 17) = data(i, cProdName)
            out(n, 18) = data(i, cCat)
            out(n, 19) = data(i, cSub)
            out(n, 20) = data(i, cChan)
            out(n, 21) = data(i, cRep)
            out(n, 22) = data(i, cPay)
            out(n, 23) = qty
            out(n, 24) = price
            out(n, 25) = disc
            out(n, 26) = gross
            out(n, 27) = discAmt
            out(n, 28) = revenue
            out(n, 29) = cogs
            out(n, 30) = profit
            If revenue <> 0 Then out(n, 31) = profit / revenue Else out(n, 31) = 0
            If KeyExists(seenOrders, orderId) Then          ' OrderFlag
                out(n, 32) = 0
            Else
                seenOrders.Add 1, orderId
                out(n, 32) = 1
            End If
            If rowStatus(i) = 1 Then out(n, 33) = "Warning" Else out(n, 33) = "OK"
        End If
    Next i
    BuildSalesRows = n
End Function
