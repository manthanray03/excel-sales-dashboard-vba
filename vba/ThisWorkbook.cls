VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================================
' ThisWorkbook
'===============================================================================
Option Explicit

Private Sub Workbook_Open()
    On Error Resume Next
    Application.StatusBar = False
    ThisWorkbook.Worksheets(SH_DASHBOARD).Activate

    ' slicers are created from code the first time the file is opened with macros on
    If ThisWorkbook.SlicerCaches.Count = 0 Then
        RefreshPivotCaches
        EnsureSlicers
    End If
End Sub
