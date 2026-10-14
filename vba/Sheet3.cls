VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================================
' Sheet3 = Data_Entry sheet
'===============================================================================
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    ' picking a product fills in its list price (can still be overwritten)
    If Intersect(Target, Me.Range("in_Product")) Is Nothing Then Exit Sub
    FillListPrice
End Sub
