Attribute VB_Name = "Module1"

Option Explicit

' ============================================================
' date: 2023-06M-07
' Module1
' ============================================================

Sub No_01_wwms_W()
Attribute No_01_wwms_W.VB_ProcData.VB_Invoke_Func = "w\n14"
' date: 2023-05M-09 by mlabrkic
' wwms zadatak, (Select All: CTRL-a), (Copy: CTRL-c)
'
' CTRL-w

  Dim i As Long
  Dim temp1 As String

  Dim iBroj_kartice As Integer, iDonatNote As Integer
  Dim iUserName As Integer, iInstallationAddress As Integer
  Dim iID_usluge As Integer, iBandwidth_usluge As Integer, iBandwidth_usluge_new As Integer

  Dim sBroj_kartice As String, sDonatNote As String, sDonatNoteFind As String
  Dim sUserName As String, sInstallationAddress As String
  Dim sID_usluge As String, sBandwidth_usluge As String, sBandwidth_usluge_new As String

  Dim iRepHrv As Integer

  ' Dim wb As Workbook
  ' Dim ws As Worksheet

  ' Set wb = ThisWorkbook
  ' Set ws = wb.Worksheets("WWMS")

  sDonatNoteFind = Worksheets("POPISI").Range("J7").Value

  Worksheets("WWMS").Activate
  Range("A:A").ClearContents  ' Clear All
  Range("A1").Select

  ' ws.PasteSpecial Format:="Unicode Text", Link:=False, DisplayAsIcon:=False
  ActiveSheet.PasteSpecial Format:="Unicode Text"

  iBroj_kartice = 0
  iDonatNote = 0
  iUserName = 0
  iInstallationAddress = 0
  iID_usluge = 0
  iBandwidth_usluge = 0
  iBandwidth_usluge_new = 0

  ' 1.red
  For i = 1 To 200
    ' temp1 = Worksheets("WWMS").Cells(i, 1).value
    temp1 = ActiveSheet.Cells(i, 1).Value

    If (iBroj_kartice = 0) Then
        iBroj_kartice = InStr(1, temp1, "Broj kartice:", vbTextCompare)
        sBroj_kartice = Trim(Mid(temp1, iBroj_kartice + 14, 9))
        If (iBroj_kartice > 0) Then
            iBroj_kartice = 1
        End If
    End If

    If (iDonatNote = 0) Then
        iDonatNote = InStr(1, temp1, sDonatNoteFind, vbTextCompare) ' sDonatNoteFind, because Crotian character
        sDonatNote = Trim(Mid(temp1, iDonatNote + 23, 1234))
        If (iDonatNote > 0) Then
            iDonatNote = 1
        End If
    End If

    If (iUserName = 0) Then
        iUserName = InStr(1, temp1, "Korisnik (1002) Ime (Naziv):", vbTextCompare)
        sUserName = Trim(Mid(temp1, iUserName + 30, 1234))
        If (iUserName > 0) Then
            iUserName = 1
        End If
    End If

    If (iInstallationAddress = 0) Then
        iInstallationAddress = InStr(1, temp1, "Adresa instalacije:", vbTextCompare)
        sInstallationAddress = Trim(Mid(temp1, iInstallationAddress + 21, 1234))
        If (iInstallationAddress > 0) Then
            iInstallationAddress = 1
        End If
    End If

    If (iID_usluge = 0) Then
        iID_usluge = InStr(1, temp1, "ID pristupa", vbTextCompare)
        sID_usluge = Trim(Mid(temp1, iID_usluge + 13, 1234))
        If (iID_usluge > 0) Then
            iID_usluge = 1
        End If
    End If

    If (iBandwidth_usluge = 0) Then
        iBandwidth_usluge = InStr(1, temp1, "Rate String", vbTextCompare)
        sBandwidth_usluge = Trim(Mid(temp1, iBandwidth_usluge + 12, 1234))
        If (iBandwidth_usluge > 0) Then
            iBandwidth_usluge = 1
        End If
    End If

    If (iBandwidth_usluge_new = 0) Then
        iBandwidth_usluge_new = InStr(1, temp1, "New rate String", vbTextCompare)
        sBandwidth_usluge_new = Trim(Mid(temp1, iBandwidth_usluge + 16, 1234))
        If (iBandwidth_usluge_new > 0) Then
            iBandwidth_usluge_new = 1
        End If
    End If

  Next i

  Worksheets("RADNA").Activate

  Range("B2").Value = sDonatNote
  Range("B3").Value = sBroj_kartice
  Range("B4").Value = sUserName

  ' Delete:  ", REPUBLIKA HRVATSKA"
  iRepHrv = InStr(1, sInstallationAddress, ", REPUBLIKA HRVATSKA", vbTextCompare)
  If (iRepHrv > 0) Then
      sInstallationAddress = Left(sInstallationAddress, iRepHrv - 1)
  End If
  Range("B5").Value = sInstallationAddress

  ActiveSheet.Range("B21").Value = sBandwidth_usluge
  If (sBandwidth_usluge_new <> "") Then
      ActiveSheet.Range("B21").Value = sBandwidth_usluge_new
  End If

  Range("B22").Value = sID_usluge

  No_01_wwms_TP_S_W

  ' Set ws = Nothing
  ' Set wb = Nothing

End Sub


Sub No_01_wwms_TP_S_W()
Attribute No_01_wwms_TP_S_W.VB_ProcData.VB_Invoke_Func = "W\n14"
  ' date: 2023-05M-09 by mlabrkic
  ' (wwms: tehnicki proces)
  ' RADNA: edit some fields
  '
  ' CTRL-SHIFT-w

  Dim sDatum As Variant
  Dim iOIB As Integer, iDoo As Integer, iDd As Integer
  Dim sUserName As String, sInstallationAddress As String
  Dim sUserName_TRpocetna As String, sUserName_file As String

  Dim iStreet As Integer, iAvenija As Integer, iCesta As Integer
  Dim iZarez1 As Integer, iZarez2 As Integer, iBlank As Integer

  Dim iKosaCrta As Integer
  Dim sStreet1 As String, sStreet2 As String, sStreet As String

  Dim sCity As String
  Dim sVrstaTR As String  ' new, 2023_05M_09

  ' Dim wb As Workbook
  Dim ws As Worksheet

  Dim sMPnumber1 As String, sMPnumber2 As String, sMPnumber As String

  ' Set wb = ThisWorkbook
  ' Set ws = wb.Worksheets("RADNA")
  Set ws = ThisWorkbook.Worksheets("RADNA")

  '   Date  ' sDatum contains the current system date. "dd.mm.yyyy"
  sDatum = Format(Date, "dd.mm.yyyy.")
  ' Worksheets("RADNA").Cells(1, 2).Value = sDatum
  ws.Range("B1").Value = sDatum

  sUserName = ws.Range("B4")
  sInstallationAddress = ws.Range("B5")

  ' --------------------------------------------------------
  iOIB = InStrRev(sUserName, " ", , vbTextCompare)
  If (iOIB > 0) Then  ' Delete:  " OIB"
    sUserName_TRpocetna = Left(sUserName, iOIB - 1)
  Else
    sUserName_TRpocetna = sUserName
  End If

  ws.Range("B8").Value = sUserName_TRpocetna  ' #TR_naslovnica_Korisnik_Naziv

  iDoo = InStrRev(sUserName_TRpocetna, " D.O.O.", , vbTextCompare)
  iDd = InStrRev(sUserName_TRpocetna, " D.D.", , vbTextCompare)
  If (iDoo > 0) Then  ' Delete:  " D.O.O."
    sUserName_file = Trim(Left(sUserName_TRpocetna, iDoo - 1))
  ElseIf (iDd > 0) Then  ' Delete:  " D.D."
    sUserName_file = Trim(Left(sUserName_TRpocetna, iDd - 1))
  Else
    sUserName_file = sUserName_TRpocetna
  End If

  ws.Range("D4").Value = sUserName_file  ' Datoteke_Korisnik_Naziv

  ' --------------------------------------------------------
  ' sStreet, sCity
  iZarez1 = InStr(1, sInstallationAddress, ",", vbTextCompare)
  sStreet = Trim(Left(sInstallationAddress, iZarez1 - 1))
  sStreet = StrConv(sStreet, vbProperCase)  ' velika pocetna slova

  iStreet = InStr(1, sStreet, "Ulica", vbTextCompare)
  If (iStreet = 1) Then
    sStreet = Replace(sStreet, "Ulica", "Ul.")
  ElseIf (iStreet > 1) Then
    sStreet = Replace(sStreet, "Ulica", "ul.")
  End If

  iAvenija = InStr(1, sStreet, "Avenija", vbTextCompare)
  If (iAvenija = 1) Then
    sStreet = Replace(sStreet, "Avenija", "Av.")
  ElseIf (iAvenija > 1) Then
    sStreet = Replace(sStreet, "Avenija", "av.")
  End If

  iCesta = InStr(1, sStreet, "Cesta", vbTextCompare)
  If (iCesta > 1) Then
    sStreet = Replace(sStreet, "Cesta", "c.")
  End If

  iKosaCrta = InStr(1, sStreet, "/", vbTextCompare)
  If (iKosaCrta > 1) Then
'    sStreet1 = Replace(sStreet, "/", " ")
    sStreet1 = Left(sStreet, iKosaCrta - 1)
    sStreet2 = Mid(sStreet, iKosaCrta + 1)
    sStreet2 = StrConv(sStreet2, vbProperCase)  ' velika pocetna slova
    sStreet = sStreet1 + " " + sStreet2
  End If

  iZarez2 = InStr(iZarez1 + 1, sInstallationAddress, ",", vbTextCompare)
  sCity = Trim(Mid(sInstallationAddress, iZarez1 + 1, iZarez2 - iZarez1 - 1))

  sVrstaTR = ws.Range("B12").Value
  iBlank = InStr(1, sCity, " ", vbTextCompare)

  If (sVrstaTR = "") Then  ' new, 2023_05M_09
    sCity = Mid(sCity, iBlank + 1)
  End If

  sCity = StrConv(sCity, vbProperCase)  ' velika pocetna slova
  If (Left(sCity, 6) = "Zagreb") Then
      sCity = "ZG"
  End If

  ws.Range("D5").Value = sStreet + ", " + sCity  ' Datoteke_Adresa_instalacije

  ' --------------------------------------------------------
  Worksheets("RADNA").Activate

  sMPnumber1 = Trim(Range("B9").Value)   ' SVK_023_23
  sMPnumber2 = Right(sMPnumber1, 6)   ' MP-SVK-023-23 ==> 023-23
  sMPnumber2 = Replace(sMPnumber2, "-", "_")  ' 023_23
  Range("B9").Value = "SVK_" + sMPnumber2

  ActiveWorkbook.Save

  Set ws = Nothing
  ' Set wb = Nothing

End Sub


' ##########################################################


' ##########################################################

Public Function GetFirstNumberLoc(ByVal s As String) As Integer
'--> Sub No_02_Copy_uredjaj_L
' https://stackoverflow.com/questions/3547744/vba-how-to-find-position-of-first-digit-in-string

  ' Sub Test_GetFirstNumericPos()
  '  Dim iPosition As Integer
  '  iPosition = GetFirstNumberLoc("ololo123")  ' OLOLO123
  '
  '  Debug.Print iPosition  ' 6
  '
  ' End Sub

  Dim i As Integer

  For i = 1 To Len(s)
    Dim currentCharacter As String
    currentCharacter = Mid(s, i, 1)

    If IsNumeric(currentCharacter) = True Then
      GetFirstNumberLoc = i
      Exit Function
    End If
  Next i
End Function


Function FindLastRow() As Long
  ' Function Copy_only_uredjaj(i)
  ' https://stackoverflow.com/questions/43631926/lastrow-and-excel-table

  ' Gives you the last cell with data in the specified row
  ' Will not work correctly if the last row is hidden

  FindLastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row

End Function


Function Copy_only_uredjaj(i)
' Macro 11.01.2021. by mlabrkic
' CTRL + SHIFT + L (small L)

' https://learn.microsoft.com/en-us/office/vba/api/excel.range.autofilter
' https://learn.microsoft.com/en-us/office/vba/api/excel.application.range

' ------------------------------------------------------------
' https://learn.microsoft.com/en-us/office/vba/api/excel.range(object)

' https://learn.microsoft.com/en-us/office/vba/api/excel.range.copy
' https://learn.microsoft.com/en-us/office/vba/api/excel.range.pastespecial

  Dim wbBook As Workbook
  Dim wsSource As Worksheet
  Dim wsDestin As Worksheet

  Dim rnSource As Range
  Dim rnDestin As Range

  Dim rnMyRange As Range
  Dim lCount As Long

  'Initialize the Excel objects
  Set wbBook = ThisWorkbook

  With wbBook
    Set wsSource = .Worksheets("MPLS")
    Set wsDestin = .Worksheets("RADNA")
  End With

  wsSource.Range("J2").Value = wsDestin.Range("A" & (25 + i), 2).Value

  'Set the destination
  With wsDestin
    ' Set rnDestin = .Cells(25 + i, 1)
    Set rnDestin = .Range("A" & (25 + i))
  End With

  ' Note: This function use the function FindLastRow
  Set rnMyRange = wsSource.Range("A5:J" & FindLastRow())

  lCount = rnMyRange.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count

  ' Following tests if any cells (other than column header) are visible _
  ' in the AutoFilter.Range
  If lCount > 1 Then
    'Firstly, remove the AutoFilter
    rnMyRange.Parent.AutoFilterMode = False

    rnMyRange.AutoFilter Field:=10, Criteria1:="=" & wsSource.Range("J2").Value

    ' rnMyRange.Parent.AutoFilter.Range.Copy
    Set rnSource = rnMyRange.Parent.AutoFilter.Range

    'Copy/paste the visible data to the new worksheet
    rnSource.Copy
    rnDestin.PasteSpecial xlPasteValues
  Else
    MsgBox "No visible data to copy"
  End If

  wsDestin.Range("B" & (26 + i)).Select

  ' Remove the AutoFilter
  rnMyRange.Parent.AutoFilterMode = False

  Set rnMyRange = Nothing
  Set rnSource = Nothing
  Set rnDestin = Nothing

  Set wsSource = Nothing
  Set wsDestin = Nothing
  Set wbBook = Nothing

End Function


Sub No_02_Copy_only_uredjaj_malo_S_L()
Attribute No_02_Copy_only_uredjaj_malo_S_L.VB_ProcData.VB_Invoke_Func = "L\n14"
  ' date: 2023-05M-09  by mlabrkic
  ' CTRL + SHIFT + l (lowercase letter L)

  Dim sSLA As String
  Dim i As Integer

  Worksheets("RADNA").Activate
  sSLA = Range("B30").Value

  If (sSLA = "") Then
      i = 0
  Else  ' SLA, Backup
      i = 7
  End If

  Copy_only_uredjaj (i)

End Sub


Function MiniProject()

  Dim sID_1 As String, sID_2 As String, sID_3 As String, sID_primar As String
  Dim sWWMS As String
  Dim sUserName_file As String, sInstallAddress_file As String, sDatum As String

  Dim sMPnumber As String
  Dim iZG As Integer

  Dim dict1 As Scripting.Dictionary  ' Note that the dictionary indexes are zero-based!
  Dim refList1 As Range, refElem1 As Range

  Dim sUsluga As String, sID_1_short As String
  Dim iID_1_first_digit As Long

  ' Drugi link (Backup, SLA, ...)
  Dim sSLA As String
  Dim sID_1_b As String, sID_2_b As String, sID_3_b As String, sID_backup As String
  Dim sWWMS_backup As String


  Worksheets("RADNA").Activate

  sMPnumber = Range("B9").Value   ' SVK_023_18

  sID_1 = Range("B22").Value
  sID_2 = Range("C22").Value
  sID_3 = Range("D22").Value

  sWWMS = Range("B3").Value
  sUserName_file = Range("D4").Value  ' Datoteke_Korisnik_Naziv
  sInstallAddress_file = Range("D5").Value  ' Datoteke_Adresa_instalacije
  sDatum = Range("B1").Value

  ' Drugi link (Backup, SLA, ...)
  sSLA = Range("B30").Value
  sID_1_b = Range("B23").Value
  sID_2_b = Range("C23").Value
  sID_3_b = Range("D23").Value
  sWWMS_backup = Range("C3").Value

  ' --------------------------------------------------------
  ' Note: This function use the function "GetFirstNumberLoc"

  ' https://stackoverflow.com/questions/3547744/vba-how-to-find-position-of-first-digit-in-string
  iID_1_first_digit = GetFirstNumberLoc(sID_1)
  sID_1_short = Left(sID_1, iID_1_first_digit - 1)

  ' Set dict1 = CreateObject("Scripting.Dictionary")
  Set dict1 = New Scripting.Dictionary
  Set refList1 = Sheets("POPISI").Range("B2:B30") 'Range of your strings in the database

  With dict1  ' copy from refList1 to dict1
    For Each refElem1 In refList1
        If Not .Exists(refElem1) And Not IsEmpty(refElem1) Then
            .Add refElem1.Value, refElem1.Offset(0, 1).Value
        End If
    Next refElem1
  End With

'    sUsluga = Dict1(Key)
  sUsluga = dict1(sID_1_short)
  Range("B20").Value = sUsluga

  ' --------------------------------------------------------
  ' ID PRIMAR ("B18"), ID BACKUP ("B19"):
  ' Excel file: Rezervacija porta ("B28" i "B35"):
  ' Naziv UR_TR datoteke i foldera ("D6 - D8")

  iZG = InStr(1, sInstallAddress_file, "ZG", vbTextCompare)
  If (iZG > 1) Then
      sInstallAddress_file = Replace(sInstallAddress_file, "ZG", "Zagreb")
  End If
  Range("D8").Value = sMPnumber + " - " + sUserName_file + ", " + sInstallAddress_file ' Naziv foldera - mrezni disk:

  If (sSLA = "") Then
      Range("D6").Value = "UR_R1_" + sMPnumber + " - " + sUserName_file + ", " + sInstallAddress_file + ", " + sWWMS + ", " + sID_1  ' Naziv UR_TR:
      Range("D7").Value = ", " + sMPnumber + ", " + sUserName_file + ", " + sInstallAddress_file + ", " + sWWMS + ", " + sID_1 ' Naziv foldera:
      If (sID_3 = "") Then
          If (sID_2 = "") Then
              sID_primar = sID_1
          Else
              sID_primar = sID_1 + ", " + sID_2
          End If
      Else
          sID_primar = sID_1 + ", " + sID_2 + ", " + sID_3
      End If
      Range("B18").Value = sID_primar  '    Primar
'        DIS, Rezervacija porta:
      Range("B28").Value = "ME- " + sID_1 + ", k_ " + sWWMS + ", " + sUserName_file + ", " + sInstallAddress_file + ", Brki" + ChrW(263) + "_" + sDatum
      Range("B29").Value = sWWMS + " " + sID_1 + " " + sUserName_file   ' Za opis porta- J
  Else  ' SLA, Backup
      Range("D6").Value = "UR_R1_" + sMPnumber + " - " + sUserName_file + ", " + sInstallAddress_file + ", " + sWWMS + ", " + sID_1_b  ' Naziv UR_TR:
      Range("D7").Value = ", SVK_" + sMPnumber + ", " + sUserName_file + ", " + sInstallAddress_file + ", " + sWWMS + ", " + sID_1_b ' Naziv foldera:
      If (sID_3_b = "") Then
          If (sID_2_b = "") Then
              sID_backup = sID_1_b
          Else
              sID_backup = sID_1_b + ", " + sID_2_b
          End If
      Else
          sID_backup = sID_1_b + ", " + sID_2_b + ", " + sID_3_b
      End If
      Range("B19").Value = sID_backup  '    Backup ID

'        DIS, Rezervacija porta, backup:
      If (sWWMS_backup = "") Then
          sWWMS_backup = sWWMS
      End If
      Range("B35").Value = "ME- " + sID_1_b + ", k_ " + sWWMS_backup + ", " + sUserName_file + ", " + sInstallAddress_file + ", Brki" + ChrW(263) + "_" + sDatum
      Range("B36").Value = sWWMS_backup + " " + sID_1_b + " " + sUserName_file   ' Za opis porta- J
  End If

  Set refList1 = Nothing
  Set dict1 = Nothing

End Function


Function Uredjaj()

  Dim sWWMS As String
  Dim sUserName_file As String, sInstallAddress_file As String, sDatum As String
  Dim iZG As Integer

  Dim sVrstaTR As String  ' new, 2023_05M_09
  Dim sSLA As String

  Worksheets("RADNA").Activate

  sWWMS = Range("B3").Value

  Range("D4").Value = Range("B8").Value
  sUserName_file = Range("D4").Value  ' Datoteke_Korisnik_Naziv
  sInstallAddress_file = Range("D5").Value  ' Datoteke_Adresa_instalacije

  sDatum = Range("B1").Value

  ' Drugi link (Backup, SLA, ...)
  sSLA = Range("B30").Value

  ' --------------------------------------------------------
  ' ID PRIMAR ("B18"), ID BACKUP ("B19"):
  ' Rezervacija porta ("B28" i "B35"):
  ' Naziv UR_TR datoteke i foldera ("D6 - D8")

  iZG = InStr(1, sInstallAddress_file, "ZG", vbTextCompare)
  If (iZG > 1) Then
      sInstallAddress_file = Replace(sInstallAddress_file, "ZG", "Zagreb")
  End If

  sVrstaTR = Range("B12").Value
  If (sSLA = "") Then
      Worksheets("RADNA").Range("B13").Value = "TR " + sUserName_file + sVrstaTR  ' #TR

      Worksheets("RADNA").Range("D6").Value = "TR " + sUserName_file + sVrstaTR + " k_ " + sWWMS  ' Naziv UR_TR:
      Worksheets("RADNA").Range("D7").Value = ", " + sUserName_file + ", " + sInstallAddress_file + ", " + sWWMS ' Naziv foldera:

  '        DIS, Rezervacija porta:
      Worksheets("RADNA").Range("B28").Value = sUserName_file + ", " + sInstallAddress_file + ", " + sWWMS + ", Brki" + ChrW(263) + ", " + sDatum
      Worksheets("RADNA").Range("B29").Value = sWWMS + " " + sUserName_file   ' Za opis porta- J
  Else  ' Backup

  End If

End Function


Sub No_02_Copy_uredjaj_L()
Attribute No_02_Copy_uredjaj_L.VB_ProcData.VB_Invoke_Func = "l\n14"
' date: 2023-05M-09 by mlabrkic
'
' CTRL + l (lowercase letter L)


  ' Drugi link (Backup, SLA, ...)
  Dim sSLA As String
  Dim sVrstaTR As String  ' new, 2023_05M_09
  Dim sModel As String, sVendor As String, sUR_ID_kratki As String
  Dim i As Integer

  Worksheets("RADNA").Activate

  ' Drugi link (Backup, SLA, ...)
  sSLA = Range("B30").Value
  sVrstaTR = Range("B12").Value

  If (sSLA = "") Then
    i = 0
    Copy_only_uredjaj (i)  ' Call Function

    ' UR copy:
    sVendor = Worksheets("RADNA").Range("E26").Value
    sModel = Worksheets("RADNA").Range("D26").Value
    sUR_ID_kratki = Worksheets("RADNA").Range("F26").Value

    Worksheets("RADNA").Range("D9").Value = Worksheets("RADNA").Range("A26").Value
    Worksheets("RADNA").Range("D10").Value = sVendor + " " + sModel + "  " + sUR_ID_kratki
    Worksheets("RADNA").Range("D11").Value = Worksheets("RADNA").Range("B26").Value
    Worksheets("RADNA").Range("D12").Value = Worksheets("RADNA").Range("B27").Value
    Worksheets("RADNA").Range("B16").Value = Worksheets("RADNA").Range("C26").Value ' #UR_kategorija
  Else  ' SLA, Backup
    i = 7
    Copy_only_uredjaj (i)  ' Call Function

    ' UR copy:
    sVendor = Worksheets("RADNA").Range("E33").Value
    sModel = Worksheets("RADNA").Range("D33").Value
    sUR_ID_kratki = Worksheets("RADNA").Range("F33").Value

    Worksheets("RADNA").Range("D13").Value = Worksheets("RADNA").Range("A33").Value
    Worksheets("RADNA").Range("D14").Value = sVendor + " " + sModel + "  " + sUR_ID_kratki
    Worksheets("RADNA").Range("D15").Value = Worksheets("RADNA").Range("B33").Value
    Worksheets("RADNA").Range("D16").Value = Worksheets("RADNA").Range("B34").Value
    Worksheets("RADNA").Range("B17").Value = Worksheets("RADNA").Range("C33").Value ' #UR_2_kategorija
  End If


  If sVrstaTR = "" Then
    MiniProject
  Else
    Uredjaj
  End If

  ' --------------------------------------------------------
'    Copy to "MP_REPORT"
  Worksheets("MP_REPORT").Range("B2").Value = Range("B1").Value
  Worksheets("MP_REPORT").Range("C2").Value = Range("B9").Value
  Worksheets("MP_REPORT").Range("D2").Value = Range("B3").Value
  Worksheets("MP_REPORT").Range("E2").Value = Worksheets("RADNA").Range("B8").Value
  Worksheets("MP_REPORT").Range("F2").Value = Worksheets("RADNA").Range("B5").Value
  ' Port i uredjaj su na pocetku od No_06_Kopiraj_u_MP_REPORT_C_R

End Sub


' ##########################################################


Sub No_03_A_Otvori_Excel_tabl_portova_T()
Attribute No_03_A_Otvori_Excel_tabl_portova_T.VB_ProcData.VB_Invoke_Func = "t\n14"
' Macro 13.01.2021. by mlabrkic
' CTRL + t

'https://stackoverflow.com/questions/19351832/copy-from-one-workbook-and-paste-into-another

  Dim bookUredjaji As Workbook

  Dim sUR_ID_kratki As String
  Dim sTempSheet As String, sTempBook As String
  Dim i As Integer

  Dim sSLA As String
  Dim iSLA As Integer
  Dim sNetworkDrive As String

  sSLA = Worksheets("RADNA").Range("B30").Value
  If (sSLA = "") Then
      iSLA = 0
  Else  ' SLA
      iSLA = 7
  End If

  sUR_ID_kratki = ActiveWorkbook.Sheets("RADNA").Cells(26 + iSLA, 6).Value

  ' sTempBook :
  For i = 2 To 1000
      sTempSheet = ActiveWorkbook.Sheets("UR_ID_POPIS").Cells(i, 1).Value
      If (sTempSheet = "") Then
          Exit Sub
      End If

      If (sTempSheet = sUR_ID_kratki) Then
          sTempBook = ActiveWorkbook.Sheets("UR_ID_POPIS").Cells(i, 2).Value
          Exit For
      End If
  Next i

'  sTempBook = "J - port_ZAGREB.xlsx"
'  ## Open workbook first:
'  Set bookUredjaji = Workbooks.Open("\\...\R1\UREDJAJI\" & sTempBook)

'  https://en.wikipedia.org/wiki/List_of_Unicode_characters
'  Latin Extended - A

  sNetworkDrive = Worksheets("POPISI").Range("J2").Value
  Set bookUredjaji = Workbooks.Open(sNetworkDrive & "\URE" & ChrW(272) & "AJI\" & sTempBook) ' Croatian character

'  Workbooks("J - port_ZAGREB.xlsx").Activate
    bookUredjaji.Sheets(sUR_ID_kratki).Activate
'  ActiveWorkbook.Save

  Set bookUredjaji = Nothing

End Sub


Sub No_03_B_Rezerviraj_port_Excel_tablica_S_T()
Attribute No_03_B_Rezerviraj_port_Excel_tablica_S_T.VB_ProcData.VB_Invoke_Func = "T\n14"
' Macro 13.01.2021. by mlabrkic
' CTRL + SHIFT + t

'https://stackoverflow.com/questions/19351832/copy-from-one-workbook-and-paste-into-another

  Dim bookUredjaji As Workbook

  Dim sUR_ID_kratki As String
  Dim sTempSheet As String, sTempBook As String
  Dim i As Integer

  Dim sRezervacija As String

  Dim sSLA As String
  Dim iSLA As Integer

  sSLA = Worksheets("RADNA").Range("B30").Value
  If (sSLA = "") Then
      iSLA = 0
  Else  ' SLA
      iSLA = 7
  End If

  sUR_ID_kratki = ActiveWorkbook.Sheets("RADNA").Cells(26 + iSLA, 6).Value
  sRezervacija = ActiveWorkbook.Sheets("RADNA").Cells(28 + iSLA, 2).Value

  ' sTempBook :
  For i = 2 To 1000
    sTempSheet = ActiveWorkbook.Sheets("UR_ID_POPIS").Cells(i, 1).Value
    If (sTempSheet = "") Then
      Exit Sub
    End If

    If (sTempSheet = sUR_ID_kratki) Then
      sTempBook = ActiveWorkbook.Sheets("UR_ID_POPIS").Cells(i, 2).Value
      Exit For
    End If
  Next i

  Set bookUredjaji = Workbooks(sTempBook)

'  ----------------------------------------------------------------------
'    sFolder_MyDoc_Trosk = Environ("USERPROFILE") & "\Documents\" & "Troskovnik.xlsm"
'    Workbooks("J - port_plan_ZAGREB.xlsx").Activate
  bookUredjaji.Sheets(sUR_ID_kratki).Activate

'  ActiveCell.Select
  ActiveCell.Value = sRezervacija

'  Select a Cell Relative to the Active Cell :
'  https://support.microsoft.com/hr-hr/help/291308/how-to-select-cells-ranges-by-using-visual-basic-procedures-in-excel
'  10: How to Select a Cell Relative to the Active Cell
'  To select a cell that is five rows below and four columns to the left of the active cell, you can use the following example:
'  ActiveCell.Offset(5, -4).Select
'
'  To select a cell that is two rows above and three columns to the right of the active cell, you can use the following example:
'  ActiveCell.Offset(-2, 3).Select

  ActiveCell.Offset(0, -2).Select

  With Selection.Interior
      .Pattern = xlSolid
      .PatternColorIndex = xlAutomatic
      .Color = 65535 ' yellow
      .TintAndShade = 0
      .PatternTintAndShade = 0
  End With

'    ActiveWorkbook.Save
  Set bookUredjaji = Nothing

End Sub


Sub No_04_A_INV_PORT_P()
Attribute No_04_A_INV_PORT_P.VB_ProcData.VB_Invoke_Func = "p\n14"
' Macro 13.01.2021. by mlabrkic
' CTRL + p

'    Application.Run ("C:\...\01_INV_PORT\Pokreni_inventory_01_port.bat")
'    x = Shell("C:\...\01_INV_PORT\Pokreni_inventory_01_port.bat", vbNormalFocus)

  Dim x1 As Variant
  Dim sPath As String
  Dim sFile As String
  Dim sUtilsFolder As String

  Dim sSLA As String

  Dim sVendor As String, sModel As String
  Dim sUR_ID As String, sPort As String
  Dim sSlot As String, sPIC As String, sMIC As String

  Dim sDescription As String
  Dim iSLA As Integer, iPos As Integer

  Dim sVendor1 As String, sVendor2 As String, sVendor3 As String

  ' lower case
  sVendor1 = Worksheets("POPISI").Range("K1").Value ' C
  sVendor2 = Worksheets("POPISI").Range("L1").Value ' J
  sVendor3 = Worksheets("POPISI").Range("M1").Value

  sUtilsFolder = Worksheets("POPISI").Range("J4").Value

  sPath = sUtilsFolder & "\01_INV_PORT\"
  sFile = "Pokreni_inventory_01_port.bat"

  sSLA = Worksheets("RADNA").Range("B30").Value
  If (sSLA = "") Then
      iSLA = 0
  Else  ' SLA
      iSLA = 7
  End If

  sVendor = Worksheets("RADNA").Cells(26 + iSLA, 5).Value ' Korisnik se povezuje na router / switch vendora
'    sModel = Worksheets("RADNA").Cells(23 + iSLA, 4).Value ' ...
  sUR_ID = Worksheets("RADNA").Cells(26 + iSLA, 2).Value
  sPort = Worksheets("RADNA").Cells(27 + iSLA, 2).Value
  sDescription = Worksheets("RADNA").Cells(29 + iSLA, 2).Value

  If (sVendor = sVendor2) Then
      iPos = InStr(3, sPort, "/", 1)    ' Pozicija prvog "/" nakon 3 karaktera
      sSlot = Mid(sPort, 4, iPos - 4)
      sPIC = Mid(sPort, iPos + 1, 1)
      If (sPIC = 0) Or (sPIC = 1) Then
          sMIC = 0
      ElseIf (sPIC = 2) Or (sPIC = 3) Then
          sMIC = 1
      End If
      sSlot = Trim(sSlot & "/" & sMIC & "/" & sPIC)
  ElseIf (sVendor = sVendor3) Then
      sSlot = "0" + Mid(sPort, 4, 1) + "-LPU"
  End If

  Worksheets("RADNA").Cells(27 + iSLA, 3).Value = sSlot
  ActiveWorkbook.Save
  x1 = Shell(sPath + sFile + " " + sUR_ID + " " + sSlot + " " + Chr(34) & sDescription & Chr(34), vbNormalFocus)


'    umjesto " staviti:    & Chr(34) &"
'    Worksheets("Totals").Cells(cellCount + 10, 5).Formula = "=COUNTIF('" & cellCount & "'!G:G," & """H"")"
'
'    vba: umjesto ; u formulu ide ,
'    =CONCATENATE(LEFT(C32;FIND("/";C32)-1);"/";C33;"/";MID(C32;FIND("/";C32)+1;1))
'    Worksheets("DIS").Cells(34, 3).Formula = "=CONCATENATE(LEFT(C32,FIND('" / "',C32)-1),'" / "',C33,'" / "',MID(C32,FIND('" / "',C32)+1,1))"
'    isTemp1_blank = InStr(20, sTemp1_trim_blank, ":", 1)    ' Pozicija ":" nakon 20 karaktera

End Sub


Sub No_04_B_INV_UI_U()
Attribute No_04_B_INV_UI_U.VB_ProcData.VB_Invoke_Func = "u\n14"
' Macro 13.01.2021. by mlabrkic
' CTRL + u (create a UI path)

'    Application.Run ("C:\...\yyy.exe")
  Dim x1 As Variant
  Dim sPath As String
  Dim sFile As String
  Dim sUtilsFolder As String

  Dim sSLA As String

  Dim sUR_ID As String, sPort As String
  Dim iSLA As Integer, iPos As Integer

  sUtilsFolder = Worksheets("POPISI").Range("J4").Value
  sPath = sUtilsFolder & "\02_INV_UI\"
  sFile = "Pokreni_inventory_02_UI.bat"

  sSLA = Worksheets("RADNA").Range("B30").Value
  If (sSLA = "") Then
      iSLA = 0
  Else  ' SLA
      iSLA = 7
  End If

'  sVendor = Worksheets("RADNA").Cells(23 + iSLA, 5).Value ' Korisnik se povezuje na router / switch vendora
  sUR_ID = Worksheets("RADNA").Cells(26 + iSLA, 2).Value
  sPort = Worksheets("RADNA").Cells(27 + iSLA, 2).Value

'    x1 = Shell("C:\...\02_INV_UI\Pokreni_inventory_02_UI.bat", vbNormalFocus)
  x1 = Shell(sPath + sFile + " " + sUR_ID + " " + sPort, vbNormalFocus)

End Sub


Sub No_04_C_INV_ACCESS_A()
Attribute No_04_C_INV_ACCESS_A.VB_ProcData.VB_Invoke_Func = "a\n14"
' Macro 13.01.2021. by mlabrkic
'CTRL + a (open the Access path)

  Dim x1 As Variant
  Dim sPath As String
  Dim sFile As String
  Dim sUtilsFolder As String

  Dim sID_tocke As String

  Dim ws As Worksheet

  sUtilsFolder = Worksheets("POPISI").Range("J4").Value
  sPath = sUtilsFolder & "\03_INV_ACCESS\"
  sFile = "Pokreni_inventory_03_access.bat"

  sID_tocke = ActiveCell.Value

'    x1 = Shell("C:\...\03_INV_ACCESS\Pokreni_inventory_03_access.bat", vbNormalFocus)
  x1 = Shell(sPath + sFile + " " + sID_tocke, vbNormalFocus)

End Sub


Sub No_04_C_INV_ACCESS_samo_najdi_S_A()
Attribute No_04_C_INV_ACCESS_samo_najdi_S_A.VB_ProcData.VB_Invoke_Func = "A\n14"
' Macro 13.01.2021. by mlabrkic
' CTRL + SHIFT + a (Net Voice ==> open the Access path)

  Dim x1 As Variant
  Dim sPath As String
  Dim sFile As String
  Dim sUtilsFolder As String

  Dim sID_tocke As String

  Dim ws As Worksheet

  sUtilsFolder = Worksheets("POPISI").Range("J4").Value
  sPath = sUtilsFolder & "\03_INV_ACCESS\"
  sFile = "Pokreni_inventory_03_access_NV.bat"

  sID_tocke = ActiveCell.Value

  x1 = Shell(sPath + sFile + " " + sID_tocke, vbNormalFocus)

End Sub


Sub No_05_Multiple_find_and_replace_D()
Attribute No_05_Multiple_find_and_replace_D.VB_ProcData.VB_Invoke_Func = "d\n14"
' date: 2023-05M-15 by mlabrkic
' Activate Word template and replace #words
' CTRL + d (docx)

' ----------------------------------------------------------
' https://stackoverflow.com/questions/16418292/open-word-from-excel
' https://stackoverflow.com/questions/18609963/to-find-and-replace-a-text-in-the-whole-document-in-ms-word-2010-including-tabl

' ----------------------------------------------------------
' Late Binding and Early Binding:
' https://learn.microsoft.com/en-us/office/vba/api/project.application

' For application-level events, register event handlers after you set Application.Visible = True.

' Using Project From Another Application:
' Early binding
' has better performance because it loads the type library at design time.

  ' Dim pjApp As MSProject.Application
  ' Set pjApp = New MSProject.Application

  ' Dim dict1 As Scripting.Dictionary
  ' Set dict1 = New Scripting.Dictionary

' ----------------------------------------------------------
  Dim WordApp As Object, WordDoc As Object

  Set WordApp = GetObject(, "Word.Application") 'Word je vec pokrenut
  WordApp.Visible = True

  '    Set WordDoc = WordApp.Documents.Open(FileName) 'Your Word document - ucitava u Word
  '    Set WordDoc = WordApp.Documents(FileName) 'Your Word document - vec je otvoren
  Set WordDoc = WordApp.Documents(1)
  '    WordDoc.Activate ' Ako je vec otvoren

  Dim sBroj_MP As String
  Dim sUR_kategorija As String, sUR_lokacija As String, sUR_port As String, sUR_vendor As String, sUR_model As String
  Dim sUR_ID As String
  Dim sSAP_sifra As String, sSAP_naziv As String, sSAP_s_MC As String, sSAP_naziv_MC As String

  Dim i As Integer
  Dim sZX_SFP As String
  Dim sLX_ZX As String, sPath As String, sNacrt As String
  Dim sTRname As String, sFName As String

  ' Drugi link (Backup, SLA, ...)
  Dim sSLA As String
  Dim sUR_2_kategorija As String, sUR_2_lokacija As String, sUR_2_port As String, sUR_2_vendor As String
  Dim sUR_2_ID As String
  Dim sSAP_2_sifra As String, sSAP_2_naziv As String

  Dim sVrstaTR As String
  Dim s10GE As String
  Dim sUPE_vendor As String

  Dim sVendor1 As String, sVendor2 As String, sVendor3 As String
  Dim sVendor1_model1 As String, sVendor1_model2 As String, sVendor1_model3 As String, sVendor1_model4 As String ' C
  Dim sVendor2_model1 As String, sVendor2_model2 As String, sVendor2_model3 As String, sVendor2_model4 As String ' J
  Dim sVendor3_model1 As String, sVendor3_model2 As String, sVendor3_model3 As String, sVendor3_model4 As String

  Dim sTR As String
  Dim Key As Variant

' ----------------------------------------------------------
' https://learn.microsoft.com/en-us/office/vba/project/concepts/ole-programmatic-identifiers-late-binding-and-early-binding-project

' Needs a reference to the Microsoft Scripting Runtime library.
' In the Tools menu, choose References to open the References - VBA Project dialog box.

  Dim dict1 As Scripting.Dictionary  ' Note that the dictionary indexes are zero-based!
  Dim refList1 As Range, refElem1 As Range

' ----------------------------------------------------------
'  "RADNA", "A1:A21" --> Dictionary
  Set dict1 = New Scripting.Dictionary
  Set refList1 = Sheets("RADNA").Range("A1:A21") 'Range of your strings in the database

  With dict1  ' copy from refList1 to dict1
    For Each refElem1 In refList1
        If Not .Exists(refElem1) And Not IsEmpty(refElem1) Then
            .Add refElem1.Value, refElem1.Offset(0, 1).Value
        End If
    Next refElem1
  End With

'ActiveDocument.Content.Find.Execute FindText:="", ReplaceWith:="", MatchWholeWord:=True, Replace:=wdReplaceAll, Wrap:=wdFindContinue, Format:=False

  For Each Key In dict1  ' Find "Key", and replace in Word
    With WordDoc.Content.Find
      .Execute FindText:=Key, ReplaceWith:=dict1(Key), MatchWholeWord:=True, Replace:=wdReplaceAll, Wrap:=wdFindContinue
    End With
  Next Key

  Set refList1 = Nothing
  Set dict1 = Nothing

' ----------------------------------------------------------
'  "RADNA", "C9:C16" --> Dictionary
  Set dict1 = New Scripting.Dictionary
  Set refList1 = Sheets("RADNA").Range("C9:C16") 'Range of your strings in the database

  With dict1  ' copy from refList1 to dict1
    For Each refElem1 In refList1
        If Not .Exists(refElem1) And Not IsEmpty(refElem1) Then
            .Add refElem1.Value, refElem1.Offset(0, 1).Value
        End If
    Next refElem1
  End With

  For Each Key In dict1  ' Find "Key", and replace in Word
    With WordDoc.Content.Find
      .Execute FindText:=Key, ReplaceWith:=dict1(Key), MatchWholeWord:=True, Replace:=wdReplaceAll, Wrap:=wdFindContinue
    End With
  Next Key

  Set refList1 = Nothing
  Set dict1 = Nothing

' ----------------------------------------------------------
  ' lower case
  sVendor1 = Worksheets("POPISI").Range("K1").Value ' C
  sVendor2 = Worksheets("POPISI").Range("L1").Value ' J
  sVendor3 = Worksheets("POPISI").Range("M1").Value

  sVendor1_model1 = Worksheets("POPISI").Range("K2").Value ' C
  sVendor1_model2 = Worksheets("POPISI").Range("K3").Value ' C
  sVendor1_model3 = Worksheets("POPISI").Range("K4").Value ' C
  sVendor1_model4 = Worksheets("POPISI").Range("K5").Value ' C

  sVendor2_model1 = Worksheets("POPISI").Range("L2").Value ' J
  sVendor2_model2 = Worksheets("POPISI").Range("L3").Value ' J
  sVendor2_model3 = Worksheets("POPISI").Range("L4").Value ' J
  sVendor2_model4 = Worksheets("POPISI").Range("L5").Value ' J

  sVendor3_model1 = Worksheets("POPISI").Range("M2").Value
  sVendor3_model2 = Worksheets("POPISI").Range("M3").Value
  sVendor3_model3 = Worksheets("POPISI").Range("M4").Value
  sVendor3_model4 = Worksheets("POPISI").Range("M5").Value

  ' UREDJAJ 1:
  sUR_model = Worksheets("RADNA").Range("D26").Value
  sUR_vendor = Worksheets("RADNA").Range("E26").Value ' Korisnik se povezuje na router / switch vendora

'  NACRT (MC, UR_1):
  sZX_SFP = Worksheets("RADNA").Range("B11").Value   ' ZX SFP (80 km) ili ER SFP+ (10GE, 40 km)
  s10GE = Worksheets("RADNA").Range("B14").Value   ' 10GE (10 Gigabit Ethernet)

  If (sZX_SFP = "ZX_SFP") Then  ' ...
      i = 1
      sLX_ZX = "_ZX"
  Else
      i = 0
      sLX_ZX = "_LX"
  End If

  If (sUR_vendor = sVendor2) Then  ' ...
      If (sUR_model = sVendor2_model1) Then ' M
          sNacrt = sVendor2 + "_" + sVendor2_model1 + sLX_ZX + ".png"
      ElseIf (sUR_model = sVendor2_model2) Or (sUR_model = sVendor2_model3) Then
          sNacrt = sVendor2 + "_ACCESS_ROUTER" + sLX_ZX + ".png"
      Else ' M
          sNacrt = sVendor2 + "_" + sVendor2_model1 + sLX_ZX + ".png"
      End If
  ElseIf (sUR_vendor = sVendor3) Then  ' ...
      If (sUR_model = sVendor3_model1) Then ' N
          sNacrt = sVendor3 + "_" + sVendor3_model1 + sLX_ZX + ".png"
      ElseIf (sUR_model = sVendor3_model2) Or (sUR_model = sVendor3_model3) Or (sUR_model = sVendor3_model4) Then
          sNacrt = sVendor3 + "_ACCESS_ROUTER" + sLX_ZX + ".png"
      Else
          sNacrt = sVendor3 + "_ACCESS_ROUTER" + sLX_ZX + ".png"
      End If
  ElseIf (sUR_vendor = sVendor1) Then  ' ...
      sNacrt = sVendor1 + "_" + sVendor1_model1 + sLX_ZX + ".png"
  End If

' ----------------------------------------------------------
'  SAP code (MC, UR_1):
  sUPE_vendor = Worksheets("RADNA").Range("B15").Value

  If (sUPE_vendor = "") Then  ' ...
      sUPE_vendor = sUR_vendor
  End If

  If (s10GE = "") Then
      sSAP_s_MC = Worksheets(sUPE_vendor).Cells(4 + i, 1).Value
      sSAP_naziv_MC = Worksheets(sUPE_vendor).Cells(4 + i, 2).Value & vbCr & Worksheets(sUPE_vendor).Cells(4 + i, 3).Value
      sSAP_sifra = Worksheets(sUR_vendor).Cells(4 + i, 1).Value
      sSAP_naziv = Worksheets(sUR_vendor).Cells(4 + i, 2).Value & vbCr & Worksheets(sUR_vendor).Cells(4 + i, 3).Value
  ElseIf (s10GE = "10GE LR (10 km)") Then
      sSAP_s_MC = Worksheets(sUPE_vendor).Cells(8, 1).Value
      sSAP_naziv_MC = Worksheets(sUPE_vendor).Cells(8, 2).Value & vbCr & Worksheets(sUPE_vendor).Cells(8, 3).Value
      sSAP_sifra = Worksheets(sUR_vendor).Cells(8, 1).Value
      sSAP_naziv = Worksheets(sUR_vendor).Cells(8, 2).Value & vbCr & Worksheets(sUR_vendor).Cells(8, 3).Value
  ElseIf (s10GE = "10GE ER (40 km)") Then
      sSAP_s_MC = Worksheets(sUPE_vendor).Cells(9, 1).Value
      sSAP_naziv_MC = Worksheets(sUPE_vendor).Cells(9, 2).Value & vbCr & Worksheets(sUPE_vendor).Cells(9, 3).Value
      sSAP_sifra = Worksheets(sUR_vendor).Cells(9, 1).Value
      sSAP_naziv = Worksheets(sUR_vendor).Cells(9, 2).Value & vbCr & Worksheets(sUR_vendor).Cells(9, 3).Value
  ElseIf (s10GE = "10GE ZR (80 km)") Then
      sSAP_s_MC = Worksheets(sUPE_vendor).Cells(10, 1).Value
      sSAP_naziv_MC = Worksheets(sUPE_vendor).Cells(10, 2).Value & vbCr & Worksheets(sUPE_vendor).Cells(10, 3).Value
      sSAP_sifra = Worksheets(sUR_vendor).Cells(10, 1).Value
      sSAP_naziv = Worksheets(sUR_vendor).Cells(10, 2).Value & vbCr & Worksheets(sUR_vendor).Cells(10, 3).Value
  End If

  WordDoc.Content.Find.Execute FindText:="#SAP_s_MC", ReplaceWith:=sSAP_s_MC, MatchWholeWord:=True, Replace:=wdReplaceOne, Wrap:=wdFindContinue
  WordDoc.Content.Find.Execute FindText:="#SAP_naziv_MC", ReplaceWith:=sSAP_naziv_MC, MatchWholeWord:=True, Replace:=wdReplaceOne, Wrap:=wdFindContinue
  WordDoc.Content.Find.Execute FindText:="#SAP_sifra", ReplaceWith:=sSAP_sifra, MatchWholeWord:=True, Replace:=wdReplaceOne, Wrap:=wdFindContinue
  WordDoc.Content.Find.Execute FindText:="#SAP_naziv", ReplaceWith:=sSAP_naziv, MatchWholeWord:=True, Replace:=wdReplaceOne, Wrap:=wdFindContinue

' ----------------------------------------------------------
  ' UREDJAJ 2 (drugi link):
  sSLA = Worksheets("RADNA").Range("B30").Value
  If (sSLA <> "") Then  ' SLA, Backup
      ' sBroj_2_kartice = Worksheets("RADNA").Range("C3").value
      sUR_2_vendor = Worksheets("RADNA").Range("E33").Value
      If (s10GE = "") Then
          sSAP_2_sifra = Worksheets(sUR_2_vendor).Cells(4 + i, 1).Value
          sSAP_2_naziv = Worksheets(sUR_2_vendor).Cells(4 + i, 2).Value & vbCr & Worksheets(sUR_2_vendor).Cells(4 + i, 3).Value
      ElseIf (s10GE = "10GE LR (10 km)") Then
          sSAP_2_sifra = Worksheets(sUR_2_vendor).Cells(8, 1).Value
          sSAP_2_naziv = Worksheets(sUR_2_vendor).Cells(8, 2).Value & vbCr & Worksheets(sUR_2_vendor).Cells(8, 3).Value
      ElseIf (s10GE = "10GE ER (40 km)") Then
          sSAP_2_sifra = Worksheets(sUR_2_vendor).Cells(9, 1).Value
          sSAP_2_naziv = Worksheets(sUR_2_vendor).Cells(9, 2).Value & vbCr & Worksheets(sUR_2_vendor).Cells(9, 3).Value
      ElseIf (s10GE = "10GE ZR (80 km)") Then
          sSAP_2_sifra = Worksheets(sUR_2_vendor).Cells(10, 1).Value
          sSAP_2_naziv = Worksheets(sUR_2_vendor).Cells(10, 2).Value & vbCr & Worksheets(sUR_2_vendor).Cells(10, 3).Value
      End If

      ' WordDoc.Content.Find.Execute FindText:="#Broj_2_kartice", ReplaceWith:=sBroj_2_kartice, MatchWholeWord:=True, Replace:=wdReplaceAll, Wrap:=wdFindContinue
      WordDoc.Content.Find.Execute FindText:="#SAP_2_sifra", ReplaceWith:=sSAP_2_sifra, MatchWholeWord:=True, Replace:=wdReplaceOne, Wrap:=wdFindContinue
      WordDoc.Content.Find.Execute FindText:="#SAP_2_naziv", ReplaceWith:=sSAP_2_naziv, MatchWholeWord:=True, Replace:=wdReplaceOne, Wrap:=wdFindContinue
  End If

' ----------------------------------------------------------
'  https://learn.microsoft.com/en-us/office/vba/api/word.headerfooter

  sVrstaTR = Worksheets("RADNA").Range("B12").Value
  sTR = Worksheets("RADNA").Range("B13").Value
  sBroj_MP = Worksheets("RADNA").Range("B9").Value

  If (sVrstaTR <> "") Then
'    Headers:
'    ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).Range.Find
    With WordDoc.Sections(1).Headers(wdHeaderFooterPrimary).Range.Find
        .Execute FindText:="#TR", ReplaceWith:=sTR
        .Execute FindText:="#Broj_MP", ReplaceWith:=sBroj_MP
    End With
  Else  ' Mini project
'  Footers:
    With WordDoc.Sections(2).Footers(1).Range.Find
        .Execute FindText:="#Broj_MP", ReplaceWith:=sBroj_MP
    End With
  End If

' ----------------------------------------------------------
'  sPath = ThisWorkbook.path & "\PREDLOSCI\Pictures_TR\"
  sPath = Environ("USERPROFILE") & "\Documents" & "\PREDLOSCI\Pictures_TR\"

'  https://docs.microsoft.com/en-us/office/vba/api/word.bookmarks
'  MsgBox ActiveDocument.Bookmarks(1).Name
'  With ActiveDocument

  With WordDoc
      If .Bookmarks.Exists("Nacrt_bookmark") Then   ' insert picture
          Dim wrdPic As Word.InlineShape
'              .Bookmarks("Nacrt_bookmark").Select
'              Set wrdPic = Selection.Range.InlineShapes.AddPicture(FileName:=sPath & sNacrt, LinkToFile:=False, SaveWithDocument:=True)
          Set wrdPic = .Bookmarks("Nacrt_bookmark").Range.InlineShapes.AddPicture(Filename:=sPath & sNacrt, LinkToFile:=False, SaveWithDocument:=True)
'              If (sUR_vendor = sVendor2) Then  ' ...
'                  wrdPic.ScaleHeight = 55
'                  wrdPic.ScaleWidth = 55
  '                wrdPic.LockAspectRatio = msoTrue
  '                wrdPic.Width = 15
'              Else
'                  wrdPic.ScaleHeight = 65
'                  wrdPic.ScaleWidth = 65
  '                wrdPic.LockAspectRatio = msoTrue
  '                wrdPic.Width = 15
'              End If

      End If
'          If .Bookmarks.Exists("Korisnik_Naziv") Then   ' insert text
'              .Bookmarks("Korisnik_Naziv").Range.Text = sUserName
''              Selection.Range.Paste
'          End If
  End With

' ----------------------------------------------------------
  ' UR TR_2023.docx"
  sTRname = Worksheets("RADNA").Range("D6").Value
  sFName = ActiveDocument.path & "\" & sTRname & ".docx"

  ActiveDocument.SaveAs sFName   ' OVO!!!
  ' ActiveDocument.Application.Quit ' (umjesto "wdDoc.Close")- zatvara i doc i app

  Set wrdPic = Nothing

  Set WordDoc = Nothing
  Set WordApp = Nothing

End Sub


Sub No_06_Kopiraj_u_MP_REPORT_R()
Attribute No_06_Kopiraj_u_MP_REPORT_R.VB_ProcData.VB_Invoke_Func = "r\n14"
' Macro 13.01.2021. by mlabrkic
' CTRL + r

'    https://stackoverflow.com/questions/19351832/copy-from-one-workbook-and-paste-into-another
  Dim x As Workbook
  Dim sFolder_MyDoc_Trosk As String
  Dim sMP_report As String
  Dim iRow As Integer
  Dim sVendor As String, sPopis_UPE_kod_KOR As String
  Dim sYear As String
  Dim sNetworkDrive As String

  Dim sVendor1 As String, sVendor2 As String, sVendor3 As String

  ' upper case
  sVendor1 = Worksheets("POPISI").Range("D10").Value
  sVendor2 = Worksheets("POPISI").Range("D11").Value ' J
  sVendor3 = Worksheets("POPISI").Range("D12").Value

  ' sYear = "2023"
  sYear = Format(Date, "yyyy")

'    Copy to "MP_REPORT"
'    Ostalo je na kraju od No_02_Copy_uredjaj_C_L
  Worksheets("MP_REPORT").Range("J2").Value = Worksheets("RADNA").Range("B27").Value
  Worksheets("MP_REPORT").Range("K2").Value = Worksheets("RADNA").Range("G26").Value

'    sFolder_MyDoc_Trosk = Environ("USERPROFILE") & "\Documents\" & "Troskovnik.xlsm"
  sVendor = ActiveWorkbook.Sheets("RADNA").Range("E26").Value

    ' sPopis_UPE_kod_KOR = ActiveWorkbook.Sheets("RADNA").Range("D15").value
  sPopis_UPE_kod_KOR = ""

  If (sVendor = sVendor2) Then ' J
      ActiveWorkbook.Sheets("MP_REPORT").Range("G2").Value = "1. " & sVendor2
  ElseIf (sVendor = sVendor3) Then
      ActiveWorkbook.Sheets("MP_REPORT").Range("G2").Value = "2. " & sVendor3
  ElseIf (sVendor = sVendor1) Then
      If (sPopis_UPE_kod_KOR = "") Then   ' UPE, lokacija UPS-a
          ActiveWorkbook.Sheets("MP_REPORT").Range("G2").Value = "3. " & sVendor1
      Else   ' UPE, lokacija korisnika
          ActiveWorkbook.Sheets("MP_REPORT").Range("G2").Value = "4. " & sVendor1 & "- elektricki"
      End If
  End If

    '  \\...\R1\UREDJAJI\MP REPORT\2019\MP_report_UR_2019.xlsx
    sMP_report = "MP_report_UR_" & sYear & ".xlsx"

  '## Open workbook first:
  sNetworkDrive = Worksheets("POPISI").Range("J2").Value
  Set x = Workbooks.Open(sNetworkDrive & "\URE" & ChrW(272) & "AJI\MP_REPORT\" & sYear & "\" & sMP_report) ' Croatian character

  Workbooks(sMP_report).Activate

  ' No_01: Select last row:
  With Application.WorksheetFunction
      Workbooks(sMP_report).Sheets("MP_REPORT_UR").Cells(Rows.Count, 2).End(xlUp).Offset(1, 0).Select
  End With

'  Workbooks("Excel_TR_radna.xlsm").Sheets("MP_REPORT").Range("B2:K2").Copy
  ThisWorkbook.Sheets("MP_REPORT").Range("B2:K2").Copy

  ActiveCell.PasteSpecial xlPasteValues

'  ActiveWorkbook.Save
  Set x = Nothing

End Sub


Sub No_07_SendEmail_M()
Attribute No_07_SendEmail_M.VB_ProcData.VB_Invoke_Func = "m\n14"
' Macro 13.01.2021. by mlabrkic
' CTRL + m
' INFO mail izvodjacima

  ' https://learn.microsoft.com/en-us/office/vba/outlook/concepts/getting-started/automating-outlook-from-other-office-applications
  ' https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/getobject-function

  ' https://learn.microsoft.com/en-us/office/vba/outlook/concepts/getting-started/automating-outlook-from-a-visual-basic-application
  ' To use early binding, you first need to set a reference to the Outlook object library.
  ' set a reference to "Microsoft Outlook xx.x" Object Library,
  ' where xx.x represents the version of Outlook that you are working with.

  'http://stackoverflow.com/questions/8994116/how-to-add-default-signature-in-outlook

  Dim signature As String

  Dim objOutlookApp As Object    ' Variable to hold reference to Microsoft Outlook.
  Dim NewMail As Outlook.MailItem

  Dim MyFont As StdFont
  Dim Subj As String, Subj1 As String, Subj2 As String, Subj3 As String
  Dim Msg1 As String, Msg2 As String

  Dim EmailAddr1 As String
  Dim EmailAddr2 As String, EmailAddr3 As String, EmailAddr4 As String, EmailAddr5 As String, EmailAddr6 As String

  Dim FnameTR As String, nazivTR As String

  Dim sID_primar As String, sID_backup As String
  Dim sSwitch_1 As String, sPort_1 As String, sSwitch_port_ID_1 As String ' Switch1
  Dim sSwitch_2 As String, sPort_2 As String, sSwitch_port_ID_2 As String ' Switch2
  Dim sBrojKartice As String
  Dim sKorisnik As String, sAdresa_instal As String

  Dim sPostovani As String, sPozdrav As String
  Dim i As Long, j As Long, wf As WorksheetFunction

  Dim sSLA As String

  sSLA = Worksheets("RADNA").Range("B30").Value

  sBrojKartice = Worksheets("RADNA").Range("B3").Value
  sKorisnik = Worksheets("RADNA").Range("B4").Value
  sAdresa_instal = Worksheets("RADNA").Range("B5").Value

  nazivTR = Worksheets("RADNA").Range("D6").Value
  FnameTR = ActiveWorkbook.path & "\" & _
    nazivTR & ".docx"

  'Switch1
  ' sTemp5 = sTemp5_1 + "&nbsp;&nbsp;&nbsp;&nbsp;" + sTemp5_2 + "&nbsp;&nbsp;&nbsp;&nbsp;" + sTemp5_3
'  i = 0
  sID_primar = Worksheets("RADNA").Range("B18").Value
  sSwitch_1 = Worksheets("RADNA").Range("G26").Value   ' 1. Switch
  sPort_1 = Worksheets("RADNA").Range("B27").Value   ' Port1
  sSwitch_port_ID_1 = sSwitch_1 + "&nbsp;&nbsp;" + "<b>" + sPort_1 + "</b>" + ",&nbsp;&nbsp;" + sID_primar    ' Switch_port_ID_primar

  If (sSLA = "") Then  ' samo primar
'      i = 0
  Else  ' SLA, Backup
'      i = 7
      'Switch2
      ' sTemp9 = sTemp9_1 + "&nbsp;&nbsp;&nbsp;&nbsp;" + sTemp9_2 + "&nbsp;&nbsp;&nbsp;&nbsp;" + sTemp9_3
      sID_backup = Worksheets("RADNA").Range("B19").Value
      sSwitch_2 = Worksheets("RADNA").Range("G33").Value   ' 2. Switch
      sPort_2 = Worksheets("RADNA").Range("B34").Value   ' Port2
      sSwitch_port_ID_2 = sSwitch_2 + "&nbsp;&nbsp;" + "<b>" + sPort_2 + "</b>" + ",&nbsp;&nbsp;" + sID_backup   ' Switch_port_ID_backup
  End If

  ' Getobject function called without the first argument,
  ' returns a reference to an instance of the application.
  Set objOutlookApp = GetObject(, "Outlook.Application") ' Outlook je vec pokrenut

  ' Create Mail Item
'  Set NewMail = Outlook.CreateItem(olMailItem)
  Set NewMail = objOutlookApp.CreateItem(olMailItem)

  ' Display the new mail
  NewMail.Display

  signature = NewMail.HTMLBody

  Subj = nazivTR + ".doc"

  EmailAddr1 = Worksheets("RADNA").Range("B10").Value  ' Author TR-a
  EmailAddr2 = Worksheets("POPISI").Range("H2").Value  ' voditelj
  EmailAddr3 = Worksheets("POPISI").Range("H3").Value  ' Alen
  EmailAddr4 = Worksheets("POPISI").Range("H4").Value  ' Robert
  EmailAddr5 = Worksheets("POPISI").Range("H5").Value  ' Boris

  EmailAddr6 = Worksheets("POPISI").Range("H6").Value  ' Mario B

  'https://stackoverflow.com/questions/28287868/choose-random-number-from-an-excel-range
  'With your data in A1 thru A4, try this macro
  Set wf = Application.WorksheetFunction
  i = wf.RandBetween(1, 3)
  j = wf.RandBetween(8, 12)
  '  MsgBox Cells(i, 11).Address & vbTab & Cells(i, 11).Value
  sPostovani = Worksheets("POPISI").Cells(i, 9).Value
  sPozdrav = Worksheets("POPISI").Cells(j, 9).Value

  'Compose message
  'Msg = "<body><font color=#ff0000>"
  'Msg = "<body><font size=2 face=Arial COLOR=red>"
  Msg1 = "<body><font size=2 face=Arial >"
  Msg1 = Msg1 & sPostovani & "<br><br>"

  Msg2 = Msg2 & "<br><br><br>" & sPozdrav & "</font></body>"


  On Error Resume Next

  With NewMail
'      .CC = EmailAddr1 & "; " & EmailAddr2 & "; " & EmailAddr3 & "; " & EmailAddr4 & "; " & EmailAddr5
      .To = EmailAddr1 & "; " & EmailAddr2 & "; " & EmailAddr3 & "; " & EmailAddr4 & "; " & EmailAddr5 & "; " & EmailAddr6
      .Subject = Subj
      ' .HTMLBody = Msg1 & sTemp5 & Msg2 & "<br><br>" & Signature

      If (sSLA = "") Then  ' samo primar
          Msg1 = Msg1 & "WWMS:  " & sBrojKartice & "<br>"
          Msg1 = Msg1 & "Korisnik:  " & sKorisnik & "<br>"
          Msg1 = Msg1 & "Adresa instalacije: " & sAdresa_instal & "<br><br>"
          Msg1 = Msg1 & "Port:  " & "<br>" & sSwitch_port_ID_1 ' & "<br><br>"
          .HTMLBody = Msg1 & Msg2 & signature
      Else  ' primar i backup
          Msg1 = Msg1 & "WWMS:  " & sBrojKartice & "<br>"
          Msg1 = Msg1 & "Korisnik:  " & sKorisnik & "<br>"
          Msg1 = Msg1 & "Adresa instalacije: " & sAdresa_instal & "<br><br>"

          Msg1 = Msg1 & "1)&nbsp;&nbsp;Prvi link <br>"
          Msg1 = Msg1 & sSwitch_port_ID_1 & "<br><br>"
          Msg1 = Msg1 & "2)&nbsp;&nbsp;Drugi link <br>"
          Msg1 = Msg1 & sSwitch_port_ID_2     ' & "<br><br>"
          .HTMLBody = Msg1 & Msg2 & signature
      End If
      '.Attachments.Add ("C:\test.txt")   'You can add files also like this
      .Attachments.Add FnameTR
      '.Send   'or use .Display
      .Save 'to Drafts folder
  End With


  On Error GoTo 0

'  PRIMJER:
'  nazivTroskovnik = Worksheets("DIS").Cells(11, 3).Value
'  FnameTroskovnik = ActiveWorkbook.Path & "\" & _
'    "OPREMA" & "_" & nazivTroskovnik & ".xlsm"
'  ActiveWorkbook.SaveAs FnameTroskovnik   ' OVO!!!

'  ActiveWorkbook.Save

  Set wf = Nothing
  Set NewMail = Nothing
  Set objOutlookApp = Nothing

End Sub


Sub No_08_Prebaci_u_novi_folder_N()
Attribute No_08_Prebaci_u_novi_folder_N.VB_ProcData.VB_Invoke_Func = "n\n14"
' Macro 15.01.2021. by mlabrkic
' CTRL + n

'https://docs.microsoft.com/en-us/office/vba/language/concepts/getting-started/calling-sub-and-function-procedures

  Dim sFolder_MyDoc As String, sFolder_godina As String

  Dim sFolder_excel_TR As String
  Dim sNaziv_excel_TR As String

  Dim sFolder_TR As String

  Dim iSubfolders_Count As Integer
  Dim sSubfolders_Count As String

  Dim sFolder_TR_MyDoc As String
  Dim sWWMS As String, sID_tocke As String

  Dim sYear As String

  ' sYear = "2023"
  sYear = Format(Date, "yyyy")

  Dim sVrstaTR As String, sFolderCount As String
  Dim sPosaoFolder As String

  Worksheets("RADNA").Activate
  sVrstaTR = Range("B12").Value

'    sFolder_MyDoc = "C:\Users\ja\Documents\"
  sFolder_MyDoc = Environ("USERPROFILE") & "\Documents\"

  sPosaoFolder = Worksheets("POPISI").Range("J5").Value

  If (sVrstaTR = "") Then
    sFolderCount = sPosaoFolder & "\TR_" & sYear
  Else
    sFolderCount = sPosaoFolder & "\AKTIVA_" & sYear
  End If

'    Call Folder_Count
  iSubfolders_Count = Folder_Count(sFolderCount)

  If (iSubfolders_Count < 10) Then
    sSubfolders_Count = "00" & iSubfolders_Count
  ElseIf ((iSubfolders_Count > 9) And (iSubfolders_Count < 100)) Then
    sSubfolders_Count = "0" & iSubfolders_Count
  Else
    sSubfolders_Count = iSubfolders_Count
  End If

  sFolder_godina = sFolderCount & "\"
  sFolder_excel_TR = Range("D7").Value

  sFolder_TR = sFolder_godina & sSubfolders_Count & sFolder_excel_TR & "\"
  sNaziv_excel_TR = Range("D6").Value & ".docx"

'    CreatePathTo ("C:\TR_2023\zz" & sFolder_moj_PC & "\nazivi_temp.txt")
'    103, SVK_SVK_164_23, T-HT, Ul. J. Broza Tita bb, Zabok, 53614744, SPM592883
  CreatePathTo (sFolder_TR)

  FileCopy sFolder_MyDoc & sNaziv_excel_TR, sFolder_TR & sNaziv_excel_TR ' Copy source to target.
  ActiveWorkbook.SaveCopyAs sFolder_TR & "Excel_TR_radna.xlsm"

  If (sVrstaTR = "") Then
'    MyDoc:
    sWWMS = Range("B3").Value
    sID_tocke = Range("B22").Value
'    sFolder_TR_MyDoc = sFolder_MyDoc & "UR_MB_" & sSubfolders_Count & "_" & sWWMS & "_" & sID_tocke & "\"
    sFolder_TR_MyDoc = sFolder_MyDoc & "UR_MB_" & sSubfolders_Count & "\"

    CreatePathTo (sFolder_TR_MyDoc)
    FileCopy sFolder_MyDoc & sNaziv_excel_TR, sFolder_TR_MyDoc & sNaziv_excel_TR
  End If

End Sub



'#################################################################################
'requires reference to Microsoft Scripting Runtime


Public Function CreatePathTo(path As String) As Boolean
'https://stackoverflow.com/questions/10803834/is-there-a-way-to-create-a-folder-and-sub-folders-in-excel-vba
'
'The following code handles both paths to a drive (like "C:\Users...") and to a server address (style: "\Server\Path.."),
'it takes a path as an argument and automatically strips any file names from it
'(use "\" at the end if it's already a directory path)
'
'and it returns false if for whatever reason the folder could not be created.
'Oh yes, it also creates sub-sub-sub-directories, if this was requested.
'answered Sep 15 '17 at 14:15
'Sascha L.

'The UBound function is used with the LBound function to determine the size of an array.
'Use the LBound function to find the lower limit of an array dimension.


  Dim sect() As String    ' path sections
  Dim reserve As Integer  ' number of path sections that should be left untouched
  Dim cPath As String     ' temp path
  Dim pos As Integer      ' position in path
  Dim lastDir As Integer  ' the last valid path length
  Dim i As Integer        ' loop var


  ' unless it all works fine, assume it didn't work:
  CreatePathTo = False

  ' trim any file name and the trailing path separator at the end:
  path = Left(path, InStrRev(path, Application.PathSeparator) - 1)

  ' split the path into directory names
  sect = Split(path, "\")


  ' what kind of path is it?
  If (UBound(sect) < 2) Then ' illegal path
      Exit Function
  ElseIf (InStr(sect(0), ":") = 2) Then
      reserve = 0 ' only drive name is reserved
  ElseIf (sect(0) = vbNullString) And (sect(1) = vbNullString) Then
      reserve = 2 ' server-path - reserve "\\Server\"
  Else ' unknown type
      Exit Function
  End If


  ' check backwards from where the path is missing:
  lastDir = -1
  For pos = UBound(sect) To reserve Step -1

      ' build the path:
      cPath = vbNullString
      For i = 0 To pos
          cPath = cPath & sect(i) & Application.PathSeparator
      Next ' i

      ' check if this path exists:
      If (Dir(cPath, vbDirectory) <> vbNullString) Then
          lastDir = pos
          Exit For
      End If

  Next ' pos


  ' create subdirectories from that point onwards:
  On Error GoTo Error01
  For pos = lastDir + 1 To UBound(sect)

      ' build the path:
      cPath = vbNullString
      For i = 0 To pos
          cPath = cPath & sect(i) & Application.PathSeparator
      Next ' i

      ' create the directory:
      MkDir cPath

  Next ' pos


  CreatePathTo = True
  Exit Function

Error01:

End Function


Function Folder_Count(sFolderCount) As Integer
'https://www.ozgrid.com/forum/forum/help-forums/excel-general/25142-count-the-number-of-subfolders-within-a-folder

  ' Dim oFSO As Object
  Dim oFSO As Scripting.FileSystemObject

  Dim folder As Object
  Dim subfolders As Object

  ' Set oFSO = CreateObject("Scripting.FileSystemObject")
  Set oFSO = New Scripting.FileSystemObject

  Set folder = oFSO.GetFolder(sFolderCount)
  Set subfolders = folder.subfolders
'  MsgBox subfolders.Count
  Folder_Count = subfolders.Count

  Set oFSO = Nothing
  Set folder = Nothing
  Set subfolders = Nothing
  'release memory

End Function


Sub ProtectActiveSheet()
' date: 2023-05M-09 by mlabrkic
'https://stackoverflow.com/questions/3037400/how-to-lock-the-data-in-a-cell-in-excel-using-vba

'You can first choose which cells you don't want to be protected (to be user-editable)
'by setting the Locked status of them to False

  Dim ws As Worksheet
  Set ws = ActiveSheet

  ws.Protect DrawingObjects:=True, Contents:=True, _
      Scenarios:=True

'  ActiveSheet.Unprotect
  Set ws = Nothing

End Sub

Sub UnprotectActiveSheet()
' date: 2023-05M-09 by mlabrkic
'https://stackoverflow.com/questions/3037400/how-to-lock-the-data-in-a-cell-in-excel-using-vba

'You can first choose which cells you don't want to be protected (to be user-editable)
'by setting the Locked status of them to False

  Dim ws As Worksheet
  Set ws = ActiveSheet

'  ws.Protect DrawingObjects:=True, Contents:=True, _
'      Scenarios:=True

  ActiveSheet.Unprotect
  Set ws = Nothing

End Sub

