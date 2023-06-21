VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

' ============================================================
' ThisWorkbook
' ============================================================

' date: 2023-05M-14 19:53:17
'https://stackoverflow.com/questions/62506679/select-specific-cell-after-pressing-tab
'==> RADNA

Private Sub Workbook_Activate()
    Application.OnKey "{TAB}", "!Sheet1.TabPress"
End Sub

Private Sub Workbook_Deactivate()
    Application.OnKey "{TAB}"
End Sub

