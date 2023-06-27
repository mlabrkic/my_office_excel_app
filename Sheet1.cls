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


Private Sub Worksheet_SelectionChange(ByVal Target As Range)
' date: 2023-06M-27 10:45:32
'https:/learn.microsoft.com/en-us/office/vba/api/excel.worksheet.selectionchange
'https://learn.microsoft.com/en-us/office/vba/api/excel.application.intersect
'https://learn.microsoft.com/en-us/office/vba/api/excel.application.selection

  ' Dim myRange1 As Range, myRange2 As Range
  Dim mySelection As Range
  Dim isect1 As Range, isect2 As Range

  'Set the range
  ' Set myRange1 = Range("B25")  ' UR primar
  ' Set myRange2 = Range("B32")  ' UR backup

  'Get the cell (or range) that the user selected
  Set mySelection = Selection

  ' Set isect1 = Application.Intersect(myRange1, mySelection)  ' UR primar
  Set isect1 = Application.Intersect(Range("B25"), mySelection)  ' UR primar, myRange1
  Set isect2 = Application.Intersect(Range("B32"), mySelection)  ' UR backup, myRange2

  'Check if the selection is inside the range "myRange1".
  If Not (isect1 Is Nothing) Then
    'Selection IS inside the range.
'    MsgBox "Selection is inside the test range."
    ComboBox1.Activate
  End If

  'Check if the selection is inside the range "myRange2".
  If Not (isect2 Is Nothing) Then
    'Selection IS inside the range.
    ComboBox2.Activate
  End If

  Set isect1 = Nothing
  Set isect2 = Nothing

  Set mySelection = Nothing
  ' Set myRange1 = Nothing
  ' Set myRange2 = Nothing

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

