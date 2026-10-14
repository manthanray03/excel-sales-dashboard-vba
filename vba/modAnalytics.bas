Attribute VB_Name = "modAnalytics"
'===============================================================================
' modAnalytics
' Business logic that is awkward to do with formulas/pivots:
'   BuildCustomerRFM  - Recency / Frequency / Monetary scores (1-5) + segment
'   BuildProductABC   - Pareto (ABC) classification of products by revenue
'   BuildCommission   - quarterly commission per rep with stepped (marginal) tiers
' All three read tblSales and fully rewrite their own output table.
'===============================================================================
Option Explicit

'------------------------------------------------------------------------------
' RFM
' R = days since last order (fewer is better), F = number of orders,
' M = total revenue. Each is scored 1-5 by where the customer sits compared with
' everyone else (share of customers strictly below them), so ties always get
' the same score.
'------------------------------------------------------------------------------
Public Sub BuildCustomerRFM()
    Dim loSales As ListObject, data As Variant, out() As Variant, n As Long
    Set loSales = RequireTable("tblSales")
    If TableRowCount(loSales) > 0 Then
        data = TableToArray(loSales)
        n = ComputeRFM(data, HeaderMap(loSales), TableRowCount(loSales), out)
    End If
    If n = 0 Then
        WriteTable RequireTable("tblRFM"), Empty, 0
    Else
        WriteTable RequireTable("tblRFM"), out, n
    End If
End Sub

' One row per customer (tblRFM columns), sorted by revenue. Returns the row count.
Public Function ComputeRFM(ByRef data As Variant, ByVal map As Collection, ByVal nSales As Long, _
                           ByRef out() As Variant) As Long
    Dim custIdx As Collection, n As Long, i As Long, k As Long
    Dim ids() As String, nm() As Variant, seg() As Variant, region() As Variant
    Dim firstD() As Date, lastD() As Date, orders() As Double, rev() As Double, prof() As Double
    Dim rec() As Double, sortedRec() As Double, sortedF() As Double, sortedM() As Double
    Dim keys() As Double, idx() As Long
    Dim asOf As Date, d As Date, maxD As Date, id As String
    Dim r As Long, f As Long, m As Long
    Dim cCust As Long, cName As Long, cSeg As Long, cRegion As Long, cOD As Long, cFlag As Long
    Dim cRev As Long, cProf As Long

    If nSales = 0 Then Exit Function
    cCust = FieldIndex(map, "CustomerID"): cName = FieldIndex(map, "CustomerName")
    cSeg = FieldIndex(map, "Segment"):     cRegion = FieldIndex(map, "Region")
    cOD = FieldIndex(map, "OrderDate"):    cFlag = FieldIndex(map, "OrderFlag")
    cRev = FieldIndex(map, "Revenue"):     cProf = FieldIndex(map, "Profit")

    ' one pass to aggregate by customer
    Set custIdx = New Collection
    ReDim ids(1 To nSales): ReDim nm(1 To nSales): ReDim seg(1 To nSales): ReDim region(1 To nSales)
    ReDim firstD(1 To nSales): ReDim lastD(1 To nSales)
    ReDim orders(1 To nSales): ReDim rev(1 To nSales): ReDim prof(1 To nSales)

    For i = 1 To nSales
        id = CStr(data(i, cCust))
        d = data(i, cOD)
        If d > maxD Then maxD = d
        If KeyExists(custIdx, id) Then
            k = custIdx.Item(id)
        Else
            n = n + 1
            k = n
            custIdx.Add k, id
            ids(k) = id: nm(k) = data(i, cName): seg(k) = data(i, cSeg): region(k) = data(i, cRegion)
            firstD(k) = d: lastD(k) = d
        End If
        If d < firstD(k) Then firstD(k) = d
        If d > lastD(k) Then lastD(k) = d
        orders(k) = orders(k) + data(i, cFlag)
        rev(k) = rev(k) + data(i, cRev)
        prof(k) = prof(k) + data(i, cProf)
    Next i

    asOf = maxD + 1                         ' "today" for the data set = day after the last sale
    ReDim rec(1 To n): ReDim sortedRec(1 To n): ReDim sortedF(1 To n): ReDim sortedM(1 To n)
    For k = 1 To n
        rec(k) = asOf - lastD(k)
        sortedRec(k) = rec(k)
        sortedF(k) = orders(k)
        sortedM(k) = rev(k)
    Next k
    QuickSortDoubles sortedRec, 1, n
    QuickSortDoubles sortedF, 1, n
    QuickSortDoubles sortedM, 1, n

    ' output sorted by revenue (high to low)
    ReDim keys(1 To n): ReDim idx(1 To n)
    For k = 1 To n
        keys(k) = -Round2(rev(k))
        idx(k) = k
    Next k
    QuickSortIndex keys, ids, idx, 1, n

    ReDim out(1 To n, 1 To 15)
    For i = 1 To n
        k = idx(i)
        r = QuintileScore(CountGreater(sortedRec, n, rec(k)), n)     ' fewer days since last order = better
        f = QuintileScore(CountLess(sortedF, n, orders(k)), n)
        m = QuintileScore(CountLess(sortedM, n, rev(k)), n)
        out(i, 1) = ids(k)
        out(i, 2) = nm(k)
        out(i, 3) = seg(k)
        out(i, 4) = region(k)
        out(i, 5) = firstD(k)
        out(i, 6) = lastD(k)
        out(i, 7) = orders(k)
        out(i, 8) = Round2(rev(k))
        out(i, 9) = Round2(prof(k))
        out(i, 10) = rec(k)
        out(i, 11) = r
        out(i, 12) = f
        out(i, 13) = m
        out(i, 14) = r * 100 + f * 10 + m                 ' e.g. 545
        out(i, 15) = RfmSegment(r, f, m)
    Next i
    ComputeRFM = n
End Function

'------------------------------------------------------------------------------
' ABC analysis: rank products by revenue, A = the products that make up the
' first 80% of revenue, B = the next 15%, C = the last 5%.
'------------------------------------------------------------------------------
Public Sub BuildProductABC()
    Dim loSales As ListObject, data As Variant, out() As Variant, n As Long
    Set loSales = RequireTable("tblSales")
    If TableRowCount(loSales) > 0 Then
        data = TableToArray(loSales)
        n = ComputeABC(data, HeaderMap(loSales), TableRowCount(loSales), out)
    End If
    If n = 0 Then
        WriteTable RequireTable("tblABC"), Empty, 0
    Else
        WriteTable RequireTable("tblABC"), out, n
    End If
End Sub

' One row per product (tblABC columns), ranked by revenue. Returns the row count.
Public Function ComputeABC(ByRef data As Variant, ByVal map As Collection, ByVal nSales As Long, _
                           ByRef out() As Variant) As Long
    Dim prodIdx As Collection, n As Long, i As Long, k As Long
    Dim ids() As String, nm() As Variant, cat() As Variant
    Dim units() As Double, rev() As Double, prof() As Double
    Dim keys() As Double, idx() As Long
    Dim total As Double, cum As Double, prevCum As Double, share As Double, rv As Double
    Dim id As String, cls As String
    Dim cProd As Long, cName As Long, cCat As Long, cQty As Long, cRev As Long, cProf As Long

    If nSales = 0 Then Exit Function
    cProd = FieldIndex(map, "ProductID"): cName = FieldIndex(map, "ProductName")
    cCat = FieldIndex(map, "Category"):   cQty = FieldIndex(map, "Quantity")
    cRev = FieldIndex(map, "Revenue"):    cProf = FieldIndex(map, "Profit")

    Set prodIdx = New Collection
    ReDim ids(1 To nSales): ReDim nm(1 To nSales): ReDim cat(1 To nSales)
    ReDim units(1 To nSales): ReDim rev(1 To nSales): ReDim prof(1 To nSales)
    For i = 1 To nSales
        id = CStr(data(i, cProd))
        If KeyExists(prodIdx, id) Then
            k = prodIdx.Item(id)
        Else
            n = n + 1
            k = n
            prodIdx.Add k, id
            ids(k) = id: nm(k) = data(i, cName): cat(k) = data(i, cCat)
        End If
        units(k) = units(k) + data(i, cQty)
        rev(k) = rev(k) + data(i, cRev)
        prof(k) = prof(k) + data(i, cProf)
    Next i

    ReDim keys(1 To n): ReDim idx(1 To n)
    For k = 1 To n
        keys(k) = -Round2(rev(k))
        idx(k) = k
    Next k
    QuickSortIndex keys, ids, idx, 1, n

    For i = 1 To n
        total = total + Round2(rev(idx(i)))
    Next i

    ReDim out(1 To n, 1 To 11)
    For i = 1 To n
        k = idx(i)
        rv = Round2(rev(k))
        If total <> 0 Then share = rv / total Else share = 0
        prevCum = cum
        cum = cum + share
        If prevCum < 0.8 Then
            cls = "A"
        ElseIf prevCum < 0.95 Then
            cls = "B"
        Else
            cls = "C"
        End If
        out(i, 1) = i
        out(i, 2) = ids(k)
        out(i, 3) = nm(k)
        out(i, 4) = cat(k)
        out(i, 5) = units(k)
        out(i, 6) = rv
        out(i, 7) = Round2(prof(k))
        If rv <> 0 Then out(i, 8) = Round2(prof(k)) / rv Else out(i, 8) = 0
        out(i, 9) = share
        out(i, 10) = cum
        out(i, 11) = cls
    Next i
    ComputeABC = n
End Function

'------------------------------------------------------------------------------
' Commission: each rep has a quarterly quota (tblReps). Revenue is paid in
' slices - e.g. 2% up to 80% of quota, 3% from 80-100%, 4% from 100-120%,
' 5% above that (tblCommissionTiers) - so hitting a new tier only raises the
' rate on the revenue above it.
'------------------------------------------------------------------------------
Public Sub BuildCommission()
    Dim loSales As ListObject, loReps As ListObject, loTiers As ListObject
    Dim out() As Variant, n As Long, data As Variant

    Set loSales = RequireTable("tblSales")
    Set loReps = RequireTable("tblReps")
    Set loTiers = RequireTable("tblCommissionTiers")
    If TableRowCount(loSales) > 0 And TableRowCount(loReps) > 0 And TableRowCount(loTiers) > 0 Then
        data = TableToArray(loSales)
        n = ComputeCommission(data, HeaderMap(loSales), TableRowCount(loSales), TableToArray(loReps), _
                              HeaderMap(loReps), TableToArray(loTiers), HeaderMap(loTiers), out)
    End If
    If n = 0 Then
        WriteTable RequireTable("tblCommission"), Empty, 0
    Else
        WriteTable RequireTable("tblCommission"), out, n
    End If
End Sub

' One row per quarter per rep (tblCommission columns). Returns the row count.
Public Function ComputeCommission(ByRef data As Variant, ByVal map As Collection, ByVal nSales As Long, _
                                  ByVal reps As Variant, ByVal rMap As Collection, _
                                  ByVal tiers As Variant, ByVal tMap As Collection, ByRef out() As Variant) As Long
    Dim tierFrom() As Double, tierRate() As Double, nTiers As Long
    Dim quarters() As String, nQ As Long, qKey As String, qSeen As Collection
    Dim revByKey As Collection, key As String
    Dim i As Long, j As Long, q As Long, n As Long, tmpD As Double, tmpR As Double, tmpS As String
    Dim rep As String, quota As Double, rv As Double, comm As Double
    Dim cRep As Long, cYear As Long, cQtr As Long, cRev As Long

    If nSales = 0 Then Exit Function

    ' tiers, sorted by starting attainment
    nTiers = UBound(tiers, 1)
    ReDim tierFrom(1 To nTiers): ReDim tierRate(1 To nTiers)
    For i = 1 To nTiers
        tierFrom(i) = tiers(i, FieldIndex(tMap, "FromAttainment"))
        tierRate(i) = tiers(i, FieldIndex(tMap, "Rate"))
    Next i
    For i = 2 To nTiers                     ' tiny list, simple insertion sort
        tmpD = tierFrom(i): tmpR = tierRate(i)
        j = i - 1
        Do While j >= 1
            If tierFrom(j) <= tmpD Then Exit Do
            tierFrom(j + 1) = tierFrom(j): tierRate(j + 1) = tierRate(j)
            j = j - 1
        Loop
        tierFrom(j + 1) = tmpD: tierRate(j + 1) = tmpR
    Next i

    ' revenue per rep per quarter
    cRep = FieldIndex(map, "SalesRep"): cYear = FieldIndex(map, "Year")
    cQtr = FieldIndex(map, "Quarter"):  cRev = FieldIndex(map, "Revenue")
    Set revByKey = New Collection
    Set qSeen = New Collection
    ReDim quarters(1 To 64)
    For i = 1 To nSales
        qKey = CStr(data(i, cYear)) & "-" & CStr(data(i, cQtr))
        If Not KeyExists(qSeen, qKey) Then
            qSeen.Add True, qKey
            nQ = nQ + 1
            If nQ > UBound(quarters) Then
                ReDim Preserve quarters(1 To nQ * 2)
            End If
            quarters(nQ) = qKey
        End If
        key = LCase$(CStr(data(i, cRep))) & "|" & qKey
        If KeyExists(revByKey, key) Then
            rv = revByKey.Item(key) + data(i, cRev)
            revByKey.Remove key
            revByKey.Add rv, key
        Else
            revByKey.Add CDbl(data(i, cRev)), key
        End If
    Next i
    For i = 2 To nQ                         ' "2024-Q1" < "2024-Q2" < ... sorts as text
        tmpS = quarters(i)
        j = i - 1
        Do While j >= 1
            If quarters(j) <= tmpS Then Exit Do
            quarters(j + 1) = quarters(j)
            j = j - 1
        Loop
        quarters(j + 1) = tmpS
    Next i

    ' one row per quarter per rep (reps in the order of tblReps)
    ReDim out(1 To nQ * UBound(reps, 1), 1 To 9)
    For q = 1 To nQ
        For i = 1 To UBound(reps, 1)
            rep = CStr(reps(i, FieldIndex(rMap, "SalesRep")))
            quota = reps(i, FieldIndex(rMap, "QuarterlyQuota"))
            key = LCase$(rep) & "|" & quarters(q)
            If KeyExists(revByKey, key) Then rv = Round2(revByKey.Item(key)) Else rv = 0
            comm = TieredCommission(rv, quota, tierFrom, tierRate, nTiers)
            n = n + 1
            out(n, 1) = CLng(Left$(quarters(q), 4))
            out(n, 2) = Mid$(quarters(q), 6)
            out(n, 3) = rep
            out(n, 4) = reps(i, FieldIndex(rMap, "Region"))
            out(n, 5) = rv
            out(n, 6) = quota
            If quota <> 0 Then out(n, 7) = rv / quota Else out(n, 7) = 0
            out(n, 8) = comm
            If rv <> 0 Then out(n, 9) = comm / rv Else out(n, 9) = 0
        Next i
    Next q
    ComputeCommission = n
End Function
