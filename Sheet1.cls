VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

' ============================================================
' date: 2023-06M-07 09:41:33
' Sheet1 = RADNA
' ============================================================

Sub TabPress()
' https://stackoverflow.com/questions/62506679/select-specific-cell-after-pressing-tab

    ' If Selection.Address = "$B$9:$B$11" Then
    If Selection.Address = "$B$10" Then
        Range("D5").Select
    ElseIf Selection.Address = "$B$21" Then
        Range("B9").Select
    Else
        Range("B21").Select
    End If
End Sub


Private Sub Worksheet_SelectionChange(ByVal Target As Excel.Range)
  Dim isect1 As Range
  Dim isect2 As Range

  Set isect1 = Application.Intersect(Target, Range("B25"))   ' UR primar
  Set isect2 = Application.Intersect(Target, Range("B32"))   ' UR backup

  ' UR primar:
  If Not isect1 Is Nothing Then
    ComboBox1.Visible = True
    ComboBox1.Activate
  Else
    ComboBox1.Visible = False
  End If

  ' UR backup:
  If Not isect2 Is Nothing Then
    ComboBox2.Visible = True
    ComboBox2.Activate
  Else
    ComboBox2.Visible = False
  End If

  ComboBox1.Visible = True
  ComboBox2.Visible = True

  Set isect1 = Nothing
  Set isect2 = Nothing

End Sub


Private Sub ComboBox1_GotFocus()
' primar:
  Me.ComboBox1.DropDown

End Sub

Private Sub ComboBox1_LostFocus()
' primar:
'  https://msdn.microsoft.com/en-us/library/microsoft.office.tools.excel.controls.combobox_events.aspx
  Dim iIndex As Integer
  Dim sValue As String

'  If the ComboBox contains text, then...
  If ComboBox1.Text <> "" Then
    iIndex = Me.ComboBox1.ListIndex
    sValue = Me.ComboBox1.List(iIndex, 0)
    ' Worksheets("RADNA").Cells(25, 2).Value = sValue
  End If

End Sub


Private Sub ComboBox2_GotFocus()
' backup:
  Me.ComboBox2.DropDown

End Sub

Private Sub ComboBox2_LostFocus()
' backup:
'  https://msdn.microsoft.com/en-us/library/microsoft.office.tools.excel.controls.combobox_events.aspx
  Dim iIndex As Integer
  Dim sValue As String

'  If the ComboBox contains text, then...
  If ComboBox2.Text <> "" Then
    iIndex = Me.ComboBox2.ListIndex
    sValue = Me.ComboBox2.List(iIndex, 0)
    ' Worksheets("RADNA").Cells(32, 2).Value = sValue
  End If

End Sub

