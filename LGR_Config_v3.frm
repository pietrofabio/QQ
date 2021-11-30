VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LGR_Config_v3 
   Caption         =   "LGR Config"
   ClientHeight    =   8700.001
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   19584
   OleObjectBlob   =   "LGR_Config_v3.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "LGR_Config_v3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_loeschen_Click()

Dim intLGR_Row As Integer

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual

Me.txt_saison_von.BackColor = rgbWhite
Me.txt_saison_bis.BackColor = rgbWhite
Me.txt_dz_1.BackColor = rgbWhite
Me.txt_ezz_ab_1.BackColor = rgbWhite
Me.txt_ezz_bis_1.BackColor = rgbWhite
Me.txt_hp_1.BackColor = rgbWhite
Me.txt_vp_1.BackColor = rgbWhite

intLGR_Row = Tabelle5.Range("O1048576").End(xlUp).row ' + 1

For i = 0 To ListBox1.ListCount
    If ListBox1.Selected(i) Then
        ListBox1.RemoveItem i
        'ThisWorkbook.Sheets("mapping").Range(Cells(i + 1, 15).Address, Cells(i + 1, 29).Address).Delete Shift:=xlUp: Exit For
        ThisWorkbook.Sheets("mapping").Range(Cells(i + 2, 15).Address, Cells(i + 2, 33).Address).ClearContents: Exit For
    End If
'With lbl_Info
'    .Caption = "Eine Zeile muss ausgewählt werden."
'End With
Next i

ActiveWorkbook.Worksheets("mapping").AutoFilter.Sort.SortFields.Clear
Tabelle5.Range("O2:AF" & intLGR_Row).Sort key1:=Tabelle5.Range("P2:P" & intLGR_Row), order1:=xlAscending, Header:=xlNo
Tabelle5.Range("O2:AF" & intLGR_Row).Sort key1:=Tabelle5.Range("O2:O" & intLGR_Row), order1:=xlAscending, Header:=xlNo

intLGR_Row = Range("O1048576").End(xlUp).row
'myarray = Tabelle5.Range("O1:AG" & intLGR_Row)
myarray = Tabelle5.Range("O2:AF" & intLGR_Row)

With Me.ListBox1
    .Clear
    .ColumnCount = 18
    .ColumnHeads = False
    .ColumnWidths = "70 Pt;80 Pt;56 Pt;72 Pt;55 Pt;36 Pt;51 Pt;72 Pt;72 Pt;77 Pt;75 Pt;36 Pt;36 Pt;36 Pt;36 Pt;36 Pt;36 Pt;36 Pt"
    .List = myarray
    .MultiSelect = 0
End With

With ThisWorkbook.Sheets("mapping")
.Range("N27").FormulaR1C1 = _
        "=IFERROR(LOOKUP(2,1/(R2C15:R" & intLGR_Row & "C15<=R23C14)/(R2C16:R" & intLGR_Row & "C16>=R23C14),R2C15:R" & intLGR_Row & "C15),"""")"
.Range("N28").FormulaR1C1 = _
        "=IFERROR(LOOKUP(2,1/(R2C15:R" & intLGR_Row & "C15<=R24C14)/(R2C16:R" & intLGR_Row & "C16>=R24C14),R2C15:R" & intLGR_Row & "C15),"""")"
.Range("N29").FormulaR1C1 = _
        "=IFERROR(LOOKUP(2,1/(R2C15:R" & intLGR_Row & "C15<=R25C14)/(R2C16:R" & intLGR_Row & "C16>=R25C14),R2C15:R" & intLGR_Row & "C15),"""")"
.Range("N30").FormulaR1C1 = _
        "=IFERROR(LOOKUP(2,1/(R2C15:R" & intLGR_Row & "C15<=R26C14)/(R2C16:R" & intLGR_Row & "C16>=R26C14),R2C15:R" & intLGR_Row & "C15),"""")"
End With

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic

End Sub

Private Sub cmd_bearbeiten_Click()

Dim intLGR_Row As Integer
Dim intBearbeiten_Row As Integer

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual

Me.txt_saison_von.BackColor = rgbWhite
Me.txt_saison_bis.BackColor = rgbWhite
Me.txt_dz_1.BackColor = rgbWhite
Me.txt_ezz_ab_1.BackColor = rgbWhite
Me.txt_ezz_bis_1.BackColor = rgbWhite
Me.txt_hp_1.BackColor = rgbWhite
Me.txt_vp_1.BackColor = rgbWhite

intLGR_Row = Tabelle5.Range("O1048576").End(xlUp).row ' + 1

For i = 0 To ListBox1.ListCount
    If ListBox1.Selected(i) Then
        intBearbeiten_Row = Tabelle5.Cells(i + 2, 15).row: Exit For
    ElseIf i = ListBox1.ListCount - 1 Then
    With lbl_Info
    .Caption = "Please select a line to edit."
    Exit Sub
    End With
    End If
Next i

'With Tabelle5
'Me.txt_saison_von.Value = .Cells(intBearbeiten_Row, 15).Value
'Me.txt_saison_bis.Value = .Cells(intBearbeiten_Row, 16).Value
'Me.txt_dz_1.Value = .Cells(intBearbeiten_Row, 17).Value
'Me.txt_ezz_ab_1.Value = .Cells(intBearbeiten_Row, 18).Value
'Me.txt_ezz_bis_1.Value = .Cells(intBearbeiten_Row, 19).Value
'Me.txt_hp_1.Value = .Cells(intBearbeiten_Row, 20).Value
'Me.txt_vp_1.Value = .Cells(intBearbeiten_Row, 21).Value
'
'Me.chk_noBud_1.Value = .Cells(intBearbeiten_Row, 22).Value
'Me.chk_Mo.Value = .Cells(intBearbeiten_Row, 23).Value
'Me.chk_Di.Value = .Cells(intBearbeiten_Row, 24).Value
'Me.chk_Mi.Value = .Cells(intBearbeiten_Row, 25).Value
'Me.chk_Do.Value = .Cells(intBearbeiten_Row, 26).Value
'Me.chk_Fr.Value = .Cells(intBearbeiten_Row, 27).Value
'Me.chk_Sa.Value = .Cells(intBearbeiten_Row, 28).Value
'Me.chk_So.Value = .Cells(intBearbeiten_Row, 29).Value
'End With

'Saison, 1/2 DZ, EZZ ab/bis, HP, VP, BUD Faktor

On Error Resume Next
Tabelle5.Range("O" & intBearbeiten_Row).Value = CDate(Me.txt_saison_von.Value)
Tabelle5.Range("P" & intBearbeiten_Row).Value = CDate(Me.txt_saison_bis.Value)
Tabelle5.Range("Q" & intBearbeiten_Row).Value = Me.txt_dz_1.Value * 1
Tabelle5.Range("R" & intBearbeiten_Row).Value = Me.txt_ezz_ab_1.Value * 1
Tabelle5.Range("S" & intBearbeiten_Row).Value = Me.txt_ezz_bis_1.Value * 1
Tabelle5.Range("T" & intBearbeiten_Row).Value = Me.txt_hp_1.Value * 1
Tabelle5.Range("U" & intBearbeiten_Row).Value = Me.txt_vp_1.Value * 1
Tabelle5.Range("V" & intBearbeiten_Row).Value = Round(Me.txt_Bel_LGR.Value * 1, 2)
Tabelle5.Range("W" & intBearbeiten_Row).Value = Me.txt_Max_LGR.Value * 1
Tabelle5.Range("X" & intBearbeiten_Row).Value = Me.txt_Min_Rate_LGR.Value * 1
Tabelle5.Range("Y" & intBearbeiten_Row).Value = Me.txt_SpielR_LGR.Value * 1

If Me.txt_Bel_LGR.Value <= 1 And Me.txt_Bel_LGR.Value >= 0 Then
Tabelle5.Range("V" & intBearbeiten_Row).Value = Round(Me.txt_Bel_LGR.Value * 100, 2)
Me.txt_Bel_LGR.Value = Round(Me.txt_Bel_LGR.Value * 100, 2)
ElseIf Me.txt_Bel_LGR.Value <= 100 And Me.txt_Bel_LGR.Value >= 0 Then
Tabelle5.Range("V" & intBearbeiten_Row).Value = Round(Me.txt_Bel_LGR.Value * 1, 2)
ElseIf Me.txt_Bel_LGR.Value > 100 And Me.txt_Bel_LGR.Value < 0 Then
Me.txt_Bel_LGR.Value = 0
MsgBox "Total Occ. % value must be between 0 and 100!"
End If

On Error GoTo 0

'Check Bud, DOW

'If Me.chk_noBud_1.Value = True Then
'    Tabelle5.Range("Z" & intBearbeiten_Row).Value = ChrW(&H2713)
'Else: Tabelle5.Range("Z" & intBearbeiten_Row).Value = ChrW(&H2717)
'End If

If Me.chk_Mo.Value = True Then
    Tabelle5.Range("Z" & intBearbeiten_Row).Value = ChrW(&H2713)
Else: Tabelle5.Range("Z" & intBearbeiten_Row).Value = ChrW(&H2717)
End If

If Me.chk_Di.Value = True Then
    Tabelle5.Range("AA" & intBearbeiten_Row).Value = ChrW(&H2713)
Else: Tabelle5.Range("AA" & intBearbeiten_Row).Value = ChrW(&H2717)
End If

If Me.chk_Mi.Value = True Then
    Tabelle5.Range("AB" & intBearbeiten_Row).Value = ChrW(&H2713)
Else: Tabelle5.Range("AB" & intBearbeiten_Row).Value = ChrW(&H2717)
End If

If Me.chk_Do.Value = True Then
    Tabelle5.Range("AC" & intBearbeiten_Row).Value = ChrW(&H2713)
Else: Tabelle5.Range("AC" & intBearbeiten_Row).Value = ChrW(&H2717)
End If

If Me.chk_Fr.Value = True Then
    Tabelle5.Range("AD" & intBearbeiten_Row).Value = ChrW(&H2713)
Else: Tabelle5.Range("AD" & intBearbeiten_Row).Value = ChrW(&H2717)
End If

If Me.chk_Sa.Value = True Then
    Tabelle5.Range("AE" & intBearbeiten_Row).Value = ChrW(&H2713)
Else: Tabelle5.Range("AE" & intBearbeiten_Row).Value = ChrW(&H2717)
End If

If Me.chk_So.Value = True Then
    Tabelle5.Range("AF" & intBearbeiten_Row).Value = ChrW(&H2713)
Else: Tabelle5.Range("AF" & intBearbeiten_Row).Value = ChrW(&H2717)
End If

ActiveWorkbook.Worksheets("mapping").AutoFilter.Sort.SortFields.Clear
Tabelle5.Range("O2:AF" & intLGR_Row).Sort key1:=Tabelle5.Range("P2:P" & intLGR_Row), order1:=xlAscending, Header:=xlNo
Tabelle5.Range("O2:AF" & intLGR_Row).Sort key1:=Tabelle5.Range("O2:O" & intLGR_Row), order1:=xlAscending, Header:=xlNo

intLGR_Row = Range("O1048576").End(xlUp).row
'myarray = Tabelle5.Range("O1:AG" & intLGR_Row)
myarray = Tabelle5.Range("O2:AF" & intLGR_Row)

With Me.ListBox1
    .Clear
    .ColumnCount = 18
    .ColumnHeads = False
    .ColumnWidths = "70 Pt;80 Pt;56 Pt;72 Pt;55 Pt;36 Pt;51 Pt;72 Pt;72 Pt;77 Pt;75 Pt;36 Pt;36 Pt;36 Pt;36 Pt;36 Pt;36 Pt;36 Pt"
    .List = myarray
    .MultiSelect = 0
End With

With Tabelle5
.Range("N27").FormulaR1C1 = _
        "=IFERROR(LOOKUP(2,1/(R2C15:R" & intLGR_Row & "C15<=R23C14)/(R2C16:R" & intLGR_Row & "C16>=R23C14),R2C15:R" & intLGR_Row & "C15),"""")"
.Range("N28").FormulaR1C1 = _
        "=IFERROR(LOOKUP(2,1/(R2C15:R" & intLGR_Row & "C15<=R24C14)/(R2C16:R" & intLGR_Row & "C16>=R24C14),R2C15:R" & intLGR_Row & "C15),"""")"
.Range("N29").FormulaR1C1 = _
        "=IFERROR(LOOKUP(2,1/(R2C15:R" & intLGR_Row & "C15<=R25C14)/(R2C16:R" & intLGR_Row & "C16>=R25C14),R2C15:R" & intLGR_Row & "C15),"""")"
.Range("N30").FormulaR1C1 = _
        "=IFERROR(LOOKUP(2,1/(R2C15:R" & intLGR_Row & "C15<=R26C14)/(R2C16:R" & intLGR_Row & "C16>=R26C14),R2C15:R" & intLGR_Row & "C15),"""")"
End With

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic

End Sub

Private Sub cmd_SelectAll_Click()

Me.chk_Mo.Value = True
Me.chk_Di.Value = True
Me.chk_Mi.Value = True
Me.chk_Do.Value = True
Me.chk_Fr.Value = True
Me.chk_Sa.Value = True
Me.chk_So.Value = True

End Sub

Private Sub cmd_speichern_Click()

Dim intLGR_Row As Integer
Dim VLookupValue As Long
Dim VLookupRng As Range
Dim Firstlogin As Boolean

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual

' ***************Datumprüfung************************************

Me.txt_saison_von.BackColor = rgbWhite
Me.txt_saison_bis.BackColor = rgbWhite
Me.txt_dz_1.BackColor = rgbWhite
Me.txt_ezz_ab_1.BackColor = rgbWhite
Me.txt_ezz_bis_1.BackColor = rgbWhite
Me.txt_hp_1.BackColor = rgbWhite
Me.txt_vp_1.BackColor = rgbWhite

If Me.txt_saison_von.Value = "" And Me.txt_saison_bis.Value = "" Then
    Me.txt_saison_von.BackColor = rgbPink
    Me.txt_saison_bis.BackColor = rgbPink
    Exit Sub
End If

If Me.txt_saison_von.Value = "" Then
    Me.txt_saison_von.BackColor = rgbPink
    Exit Sub
End If

If Me.txt_saison_bis.Value = "" Then
    Me.txt_saison_bis.BackColor = rgbPink
    Exit Sub
End If

With Me.txt_saison_von
    If CStr(Len(Me.txt_saison_von)) = "6" Then
        Tag = Left(Me.txt_saison_von.Value, 2)
        Monat = Mid(Me.txt_saison_von.Value, 3, 2)
        Jahr = Right(Me.txt_saison_von, 2)
        Me.txt_saison_von = DateSerial(CInt(Jahr), CInt(Monat), CInt(Tag))
        .SetFocus
    ElseIf CStr(Len(Me.txt_saison_von)) = "8" And Mid(Me.txt_saison_von.Value, 3, 1) = "." Then
        Tag = Left(Me.txt_saison_von.Value, 2)
        Monat = Mid(Me.txt_saison_von.Value, 4, 2)
        Jahr = Right(Me.txt_saison_von, 2)
        Me.txt_saison_von = DateSerial(CInt(Jahr), CInt(Monat), CInt(Tag))
        .SetFocus
    ElseIf CStr(Len(Me.txt_saison_von)) = "8" And Mid(Me.txt_saison_von.Value, 3, 1) <> "." Then
        Tag = Left(Me.txt_saison_von.Value, 2)
        Monat = Mid(Me.txt_saison_von.Value, 3, 2)
        Jahr = Right(Me.txt_saison_von, 4)
        Me.txt_saison_von = DateSerial(CInt(Jahr), CInt(Monat), CInt(Tag))
        .SetFocus
    ElseIf Not IsDate(.Text) Then
        MsgBox "Please enter a <Season start> date. (Dateformat: DDMMYY, DDMMYYYY, DD.MM.YYYY, oder DD/MM/YYYY)"
        '.Value = Date + 30
        .SetFocus
        Exit Sub
    End If
End With

With Me.txt_saison_bis
    If CStr(Len(Me.txt_saison_bis)) = "6" Then
        Tag = Left(Me.txt_saison_bis.Value, 2)
        Monat = Mid(Me.txt_saison_bis.Value, 3, 2)
        Jahr = Right(Me.txt_saison_bis, 2)
        Me.txt_saison_bis = DateSerial(CInt(Jahr), CInt(Monat), CInt(Tag))
        .SetFocus
    ElseIf CStr(Len(Me.txt_saison_bis)) = "8" And Mid(Me.txt_saison_von.Value, 3, 1) = "." Then
        Tag = Left(Me.txt_saison_bis.Value, 2)
        Monat = Mid(Me.txt_saison_bis.Value, 4, 2)
        Jahr = Right(Me.txt_saison_bis, 2)
        Me.txt_saison_bis = DateSerial(CInt(Jahr), CInt(Monat), CInt(Tag))
        .SetFocus
    ElseIf CStr(Len(Me.txt_saison_bis)) = "8" And Mid(Me.txt_saison_von.Value, 3, 1) <> "." Then
        Tag = Left(Me.txt_saison_bis.Value, 2)
        Monat = Mid(Me.txt_saison_bis.Value, 3, 2)
        Jahr = Right(Me.txt_saison_bis, 4)
        Me.txt_saison_bis = DateSerial(CInt(Jahr), CInt(Monat), CInt(Tag))
        .SetFocus
    ElseIf Not IsDate(.Text) Then
        MsgBox "Please enter a <Season end> date. (Dateformat: DDMMYY, DDMMYYYY, DD.MM.YYYY, oder DD/MM/YYYY)"
        '.Value = Date + 31
        .SetFocus
        Exit Sub
    End If
End With

If CDate(Me.txt_saison_von.Value) < DateSerial(2021, 1, 1) Or CDate(Me.txt_saison_von.Value) > DateSerial(2024, 12, 31) Then
    MsgBox "The <Season start> date must be between 01.01.2021 and 31.12.2024"
    Me.txt_saison_von.SetFocus
    Exit Sub
End If

If CDate(Me.txt_saison_bis.Value) < DateSerial(2021, 1, 2) Or CDate(Me.txt_saison_bis.Value) > DateSerial(2025, 1, 1) Then
    MsgBox "The <Season end> date must be between 02.01.2021 and 01.01.2025"
    Me.txt_saison_von.SetFocus
    Exit Sub
End If

If Not CDate(Me.txt_saison_von.Value) < CDate(Me.txt_saison_bis.Value) Then
    MsgBox "<Season start> date must be earlier than <Season end>."
    Me.txt_saison_von.Value = ""
    Me.txt_saison_bis.Value = ""
    Exit Sub
End If

'****************Datenprüfung************************************

If Me.txt_dz_1.Value = "" Then
    Me.txt_dz_1.BackColor = rgbPink
    Exit Sub
End If

If Me.txt_ezz_ab_1.Value = "" Then
    Me.txt_ezz_ab_1.BackColor = rgbPink
    Exit Sub
End If

If Me.txt_ezz_bis_1.Value = "" Then
    Me.txt_ezz_bis_1.BackColor = rgbPink
    Exit Sub
End If

If Me.txt_hp_1.Value = "" Then
    Me.txt_hp_1.BackColor = rgbPink
    Exit Sub
End If

If Me.txt_dz_1.Value = "" Then
    Me.txt_dz_1.BackColor = rgbPink
    Exit Sub
End If

'****************Dateneingabe************************************

'If Tabelle5.Range("O2") <> "" Then

intLGR_Row = Tabelle5.Range("O1048576").End(xlUp).row + 1

'    Else: intLGR_Row = 2
'    Firstlogin = True
'End If

'Saison, 1/2 DZ, EZZ ab/bis, HP, VP, BUD Faktor

Tabelle5.Range("O" & intLGR_Row).Value = CDate(Me.txt_saison_von.Value)
Tabelle5.Range("P" & intLGR_Row).Value = CDate(Me.txt_saison_bis.Value)
Tabelle5.Range("Q" & intLGR_Row).Value = Me.txt_dz_1.Value * 1
Tabelle5.Range("R" & intLGR_Row).Value = Me.txt_ezz_ab_1.Value * 1
Tabelle5.Range("S" & intLGR_Row).Value = Me.txt_ezz_bis_1.Value * 1
Tabelle5.Range("T" & intLGR_Row).Value = Me.txt_hp_1.Value * 1
Tabelle5.Range("U" & intLGR_Row).Value = Me.txt_vp_1.Value * 1
If Me.txt_Bel_LGR.Value <= 1 And Me.txt_Bel_LGR.Value >= 0 Then
Tabelle5.Range("V" & intLGR_Row).Value = Round(Me.txt_Bel_LGR.Value * 100, 2)
ElseIf Me.txt_Bel_LGR.Value <= 100 And Me.txt_Bel_LGR.Value >= 0 Then
Tabelle5.Range("V" & intLGR_Row).Value = Round(Me.txt_Bel_LGR.Value * 1, 2)
ElseIf Me.txt_Bel_LGR.Value > 100 And Me.txt_Bel_LGR.Value < 0 Then
Me.txt_Bel_LGR.Value = 0
MsgBox "Total Occ. % value must be between 0 and 100!"
End If
Tabelle5.Range("W" & intLGR_Row).Value = Me.txt_Max_LGR.Value * 1
Tabelle5.Range("X" & intLGR_Row).Value = Me.txt_Min_Rate_LGR.Value * 1
Tabelle5.Range("Y" & intLGR_Row).Value = Me.txt_SpielR_LGR.Value * 1

'Check Bud, DOW

'If Me.chk_noBud_1.Value = True Then
'    Tabelle5.Range("Z" & intLGR_Row).Value = ChrW(&H2713)
'Else: Tabelle5.Range("Z" & intLGR_Row).Value = ChrW(&H2717)
'End If

If Me.chk_Mo.Value = True Then
    Tabelle5.Range("Z" & intLGR_Row).Value = ChrW(&H2713)
Else: Tabelle5.Range("Z" & intLGR_Row).Value = ChrW(&H2717)
End If

If Me.chk_Di.Value = True Then
    Tabelle5.Range("AA" & intLGR_Row).Value = ChrW(&H2713)
Else: Tabelle5.Range("AA" & intLGR_Row).Value = ChrW(&H2717)
End If

If Me.chk_Mi.Value = True Then
    Tabelle5.Range("AB" & intLGR_Row).Value = ChrW(&H2713)
Else: Tabelle5.Range("AB" & intLGR_Row).Value = ChrW(&H2717)
End If

If Me.chk_Do.Value = True Then
    Tabelle5.Range("AC" & intLGR_Row).Value = ChrW(&H2713)
Else: Tabelle5.Range("AC" & intLGR_Row).Value = ChrW(&H2717)
End If

If Me.chk_Fr.Value = True Then
    Tabelle5.Range("AD" & intLGR_Row).Value = ChrW(&H2713)
Else: Tabelle5.Range("AD" & intLGR_Row).Value = ChrW(&H2717)
End If

If Me.chk_Sa.Value = True Then
    Tabelle5.Range("AE" & intLGR_Row).Value = ChrW(&H2713)
Else: Tabelle5.Range("AE" & intLGR_Row).Value = ChrW(&H2717)
End If

If Me.chk_So.Value = True Then
    Tabelle5.Range("AF" & intLGR_Row).Value = ChrW(&H2713)
Else: Tabelle5.Range("AF" & intLGR_Row).Value = ChrW(&H2717)
End If

'Max. Anzahl #, X-Bett, Gepäck

'Tabelle5.Range("N2").Value = Me.txt_bud_faktor_1.Value * 1
Tabelle5.Range("N3").Value = Me.txt_XtraBed.Value * 1
Tabelle5.Range("N4").Value = Me.txt_Luggage.Value * 1

'****************Overlapping************************************

Dim start1 As Variant, finish1 As Variant, start2 As Variant, finish2 As Variant
Dim cell As Range

intLGR_Row = Tabelle5.Range("O1048576").End(xlUp).row

If intLGR_Row <> 2 Then

start2 = Tabelle5.Range("O" & intLGR_Row).Value
finish2 = Tabelle5.Range("P" & intLGR_Row).Value

Set cell = Tabelle5.Range("O2")

For Each cell In Tabelle5.Range("O2:O" & intLGR_Row - 1)

start1 = Tabelle5.Cells(cell.row, 15).Value
finish1 = Tabelle5.Cells(cell.row, 16).Value

'If Firstlogin = True Then: Exit For

If OVERLAPS(start1, finish1, start2, finish2) Then

For i = 0 To 6
    If Tabelle5.Cells(cell.row, 26 + i).Value = ChrW(&H2713) And Tabelle5.Cells(intLGR_Row, 26 + i).Value = ChrW(&H2713) Then
    MsgBox "The season data is already entered or overlapping." & vbCrLf & "Please validate your inputs in order to avoid double bookings!", vbInformation ': Exit For
    ThisWorkbook.Sheets("mapping").Range(Cells(intLGR_Row, 15).Address, Cells(intLGR_Row, 32).Address).Delete Shift:=xlUp
    GoTo Weiter
    End If
Next i
      
End If

Next cell

End If

'****************Filtern************************************

ActiveWorkbook.Worksheets("mapping").AutoFilter.Sort.SortFields.Clear
ThisWorkbook.Worksheets("mapping").Range("O2:AF" & intLGR_Row).Sort key1:=ThisWorkbook.Worksheets("mapping").Range("P2:P" & intLGR_Row), order1:=xlAscending, Header:=xlNo
ThisWorkbook.Worksheets("mapping").Range("O2:AF" & intLGR_Row).Sort key1:=ThisWorkbook.Worksheets("mapping").Range("O2:O" & intLGR_Row), order1:=xlAscending, Header:=xlNo

'****************Visualisieren************************************

'myarray = Tabelle5.Range("O1:AG" & intLGR_Row)
myarray = Tabelle5.Range("O2:AF" & intLGR_Row)

With Me.ListBox1
    .Clear
    .ColumnCount = 18
    .ColumnHeads = False
    .ColumnWidths = "70 Pt;80 Pt;56 Pt;72 Pt;55 Pt;36 Pt;51 Pt;72 Pt;72 Pt;77 Pt;75 Pt;36 Pt;36 Pt;36 Pt;36 Pt;36 Pt;36 Pt;36 Pt"
    .List = myarray
    .MultiSelect = 0
End With

With Tabelle5
.Range("N27").FormulaR1C1 = _
        "=IFERROR(LOOKUP(2,1/(R2C15:R" & intLGR_Row & "C15<=R23C14)/(R2C16:R" & intLGR_Row & "C16>=R23C14),R2C15:R" & intLGR_Row & "C15),"""")"
.Range("N28").FormulaR1C1 = _
        "=IFERROR(LOOKUP(2,1/(R2C15:R" & intLGR_Row & "C15<=R24C14)/(R2C16:R" & intLGR_Row & "C16>=R24C14),R2C15:R" & intLGR_Row & "C15),"""")"
.Range("N29").FormulaR1C1 = _
        "=IFERROR(LOOKUP(2,1/(R2C15:R" & intLGR_Row & "C15<=R25C14)/(R2C16:R" & intLGR_Row & "C16>=R25C14),R2C15:R" & intLGR_Row & "C15),"""")"
.Range("N30").FormulaR1C1 = _
        "=IFERROR(LOOKUP(2,1/(R2C15:R" & intLGR_Row & "C15<=R26C14)/(R2C16:R" & intLGR_Row & "C16>=R26C14),R2C15:R" & intLGR_Row & "C15),"""")"
End With

Weiter:
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic

End Sub

Private Sub cmd_Schließen_Click()

Unload Me

End Sub

Private Sub cmd_UnselectAll_Click()

Me.chk_Mo.Value = False
Me.chk_Di.Value = False
Me.chk_Mi.Value = False
Me.chk_Do.Value = False
Me.chk_Fr.Value = False
Me.chk_Sa.Value = False
Me.chk_So.Value = False

End Sub

Private Sub cmd_Weekdays_Click()

Me.chk_Mo.Value = True
Me.chk_Di.Value = True
Me.chk_Mi.Value = True
Me.chk_Do.Value = True
Me.chk_Fr.Value = False
Me.chk_Sa.Value = False
Me.chk_So.Value = False

End Sub

Private Sub cmd_Weekends_Click()

Me.chk_Mo.Value = False
Me.chk_Di.Value = False
Me.chk_Mi.Value = False
Me.chk_Do.Value = False
Me.chk_Fr.Value = True
Me.chk_Sa.Value = True
Me.chk_So.Value = True

End Sub

Private Sub ListBox1_Click()

Dim intLGR_Row As Integer
Dim intBearbeiten_Row As Integer
Dim strkreuz As String

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual

Me.txt_saison_von.BackColor = rgbWhite
Me.txt_saison_bis.BackColor = rgbWhite
Me.txt_dz_1.BackColor = rgbWhite
Me.txt_ezz_ab_1.BackColor = rgbWhite
Me.txt_ezz_bis_1.BackColor = rgbWhite
Me.txt_hp_1.BackColor = rgbWhite
Me.txt_vp_1.BackColor = rgbWhite

intLGR_Row = Tabelle5.Range("O1048576").End(xlUp).row + 1

For i = 0 To ListBox1.ListCount
    If ListBox1.Selected(i) Then
        intBearbeiten_Row = Tabelle5.Cells(i + 2, 15).row
        With lbl_Info
        .Caption = ""
        End With: Exit For
    ElseIf i = ListBox1.ListCount - 1 Then
    With lbl_Info
    .Caption = "Please select a row to edit."
    End With
    Exit Sub
    End If
Next i

With Tabelle5
Me.txt_saison_von.Value = .Cells(intBearbeiten_Row, 15).Value
Me.txt_saison_bis.Value = .Cells(intBearbeiten_Row, 16).Value
Me.txt_dz_1.Value = .Cells(intBearbeiten_Row, 17).Value
Me.txt_ezz_ab_1.Value = .Cells(intBearbeiten_Row, 18).Value
Me.txt_ezz_bis_1.Value = .Cells(intBearbeiten_Row, 19).Value
Me.txt_hp_1.Value = .Cells(intBearbeiten_Row, 20).Value
Me.txt_vp_1.Value = .Cells(intBearbeiten_Row, 21).Value
Me.txt_Bel_LGR.Value = .Cells(intBearbeiten_Row, 22).Value
Me.txt_Max_LGR.Value = .Cells(intBearbeiten_Row, 23).Value
Me.txt_Min_Rate_LGR = .Cells(intBearbeiten_Row, 24).Value
Me.txt_SpielR_LGR.Value = .Cells(intBearbeiten_Row, 25).Value

'strkreuz = .Cells(intBearbeiten_Row, 26).Value
'Me.chk_noBud_1.Value = Kreuz(strkreuz)
strkreuz = .Cells(intBearbeiten_Row, 26).Value
Me.chk_Mo.Value = Kreuz(strkreuz)
strkreuz = .Cells(intBearbeiten_Row, 27).Value
Me.chk_Di.Value = Kreuz(strkreuz)
strkreuz = .Cells(intBearbeiten_Row, 28).Value
Me.chk_Mi.Value = Kreuz(strkreuz)
strkreuz = .Cells(intBearbeiten_Row, 29).Value
Me.chk_Do.Value = Kreuz(strkreuz)
strkreuz = .Cells(intBearbeiten_Row, 30).Value
Me.chk_Fr.Value = Kreuz(strkreuz)
strkreuz = .Cells(intBearbeiten_Row, 31).Value
Me.chk_Sa.Value = Kreuz(strkreuz)
strkreuz = .Cells(intBearbeiten_Row, 32).Value
Me.chk_So.Value = Kreuz(strkreuz)

End With

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic

End Sub

Private Sub UserForm_Initialize()

With Me
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 10
.Caption = "Leisure Group Configuration"
End With

'_____Textbox_____________________________________________________________________________

With Me.txt_saison_von
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 10
.TextAlign = fmTextAlignCenter
'.Value = Tabelle5.Range("R2").Value
End With

With Me.txt_saison_bis
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 10
.TextAlign = fmTextAlignCenter
'.Value = Tabelle5.Range("R3").Value
End With

With Me.txt_dz_1
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 10
.TextAlign = fmTextAlignCenter
'.Value = Tabelle5.Range("R4").Value
End With

With Me.txt_ezz_ab_1
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 10
.TextAlign = fmTextAlignCenter
'.Value = Tabelle5.Range("R5").Value
End With

With Me.txt_ezz_bis_1
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 10
.TextAlign = fmTextAlignCenter
'.Value = Tabelle5.Range("R6").Value
End With

With Me.txt_hp_1
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 10
.TextAlign = fmTextAlignCenter
'.Value = Tabelle5.Range("R7").Value
End With

With Me.txt_vp_1
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 10
.TextAlign = fmTextAlignCenter
'.Value = Tabelle5.Range("R8").Value
End With

With Me.txt_Bel_LGR
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 10
.TextAlign = fmTextAlignCenter
End With

'With Me.txt_bud_faktor_1
'.BackColor = RGB(255, 255, 255)
'.ForeColor = RGB(0, 0, 0)
'.Font.Name = "Calibri"
'.Font.Size = 10
'.TextAlign = fmTextAlignCenter
'.Value = Tabelle5.Range("N2").Value
'End With

With Me.txt_Max_LGR
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 10
.TextAlign = fmTextAlignCenter
'.Value = Tabelle5.Range("AF2").Value
End With

With Me.txt_XtraBed
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 10
.TextAlign = fmTextAlignCenter
.Value = Tabelle5.Range("N3").Value
End With

With Me.txt_Luggage
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 10
.TextAlign = fmTextAlignCenter
.Value = Tabelle5.Range("N4").Value
End With

With Me.txt_SpielR_LGR
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 10
.TextAlign = fmTextAlignCenter
'.Value = Tabelle5.Range("L6").Value
End With

With Me.txt_Min_Rate_LGR
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 10
.TextAlign = fmTextAlignCenter
'.Value = Tabelle5.Range("N4").Value
End With


'_____________Label_____________________________

With fr_1
.BackColor = RGB(255, 255, 255)
.Font.Name = "Calibri"
End With

With fr_2
.BackColor = RGB(255, 255, 255)
.Font.Name = "Calibri"
End With

With fr_3
.BackColor = RGB(255, 255, 255)
.Font.Name = "Calibri"
End With

With Me.lbl_saison_von
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 9
End With

With Me.lbl_saison_bis
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 9
End With

With Me.lbl_dz_pax
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 9
.TextAlign = fmTextAlignCenter
End With

With Me.lbl_ezz_ab
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 9
.TextAlign = fmTextAlignCenter
End With

With Me.lbl_ezz_bis
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 9
.TextAlign = fmTextAlignCenter
End With

With Me.lbl_hp
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 9
.TextAlign = fmTextAlignCenter
End With

With Me.lbl_vp
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 9
.TextAlign = fmTextAlignCenter
End With

With Me.lbl_Total_Occ_LGR
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 9
.TextAlign = fmTextAlignCenter
End With

'With Me.lbl_Budfactor
'.BackColor = RGB(255, 255, 255)
'.ForeColor = RGB(0, 0, 0)
'.Font.Name = "Calibri"
'.Font.Size = 9
'.TextAlign = fmTextAlignCenter
'End With

With Me.lbl_Max_Zi_LGR
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 9
.TextAlign = fmTextAlignCenter
End With

With Me.lbl_Extrab
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 9
.TextAlign = fmTextAlignCenter
End With

With Me.lbl_Luggage
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 9
.TextAlign = fmTextAlignCenter
End With

With Me.lbl_SpielR_LGR
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 9
.TextAlign = fmTextAlignCenter
End With

With Me.lbl_Min_Rate_LGR
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 9
.TextAlign = fmTextAlignCenter
End With

'With Me.lbl_noBud
'.BackColor = RGB(255, 255, 255)
'.ForeColor = RGB(0, 0, 0)
'.Font.Name = "Calibri"
'.Font.Size = 9
'.TextAlign = fmTextAlignCenter
'End With

With Me.lbl_Mo
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 9
.TextAlign = fmTextAlignCenter
End With

With Me.lbl_Di
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 9
.TextAlign = fmTextAlignCenter
End With

With Me.lbl_Mi
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 9
.TextAlign = fmTextAlignCenter
End With

With Me.lbl_Do
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 9
.TextAlign = fmTextAlignCenter
End With

With Me.lbl_Fr
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 9
.TextAlign = fmTextAlignCenter
End With

With Me.lbl_Sa
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 9
.TextAlign = fmTextAlignCenter
End With

With Me.lbl_So
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 9
.TextAlign = fmTextAlignCenter
End With

With Me.lbl_Info
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 9
.Caption = ""
End With

With Me.lbl_SaisonStart
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 9
.TextAlign = fmTextAlignCenter
End With

With Me.lbl_SaisonEnd
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 9
.TextAlign = fmTextAlignCenter
End With

With Me.lbl_DR
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 9
.TextAlign = fmTextAlignCenter
End With

With Me.lbl_SR20pl
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 9
.TextAlign = fmTextAlignCenter
End With

With Me.lbl_SR20mi
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 9
.TextAlign = fmTextAlignCenter
End With

With Me.lbl_HB
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 9
.TextAlign = fmTextAlignCenter
End With

With Me.lbl_FB
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 9
.TextAlign = fmTextAlignCenter
End With

With Me.lbl_TtlOcc
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 9
.TextAlign = fmTextAlignCenter
End With

With Me.lbl_MaxRooms
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 9
.TextAlign = fmTextAlignCenter
End With

With Me.lbl_MinRate
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 9
.TextAlign = fmTextAlignCenter
End With

With Me.lbl_RoomForNeg_LGR
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 9
.TextAlign = fmTextAlignCenter
End With

'With Me.lbl_Budgetiert_LGR
'.BackColor = RGB(255, 255, 255)
'.ForeColor = RGB(0, 0, 0)
'.Font.Name = "Calibri"
'.Font.Size = 9
'.TextAlign = fmTextAlignCenter
'End With

With Me.lbl_Mon_LGR
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 9
.TextAlign = fmTextAlignCenter
End With

With Me.lbl_Tue_LGR
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 9
.TextAlign = fmTextAlignCenter
End With

With Me.lbl_Wed_LGR
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 9
.TextAlign = fmTextAlignCenter
End With

With Me.lbl_Thu_LGR
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 9
.TextAlign = fmTextAlignCenter
End With

With Me.lbl_Fri_LGR
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 9
.TextAlign = fmTextAlignCenter
End With

With Me.lbl_Sat_LGR
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 9
.TextAlign = fmTextAlignCenter
End With

With Me.lbl_Sun_LGR
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 9
.TextAlign = fmTextAlignCenter
End With

'_____________Command Button_____________________________

With Me.cmd_SelectAll
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 9
End With

With Me.cmd_UnselectAll
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 9
End With

With Me.cmd_Weekdays
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 9
End With

With Me.cmd_Weekends
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 9
End With

With Me.cmd_Speichern
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 10
End With

With Me.cmd_loeschen
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 10
End With

With Me.cmd_bearbeiten
.BackColor = RGB(175, 238, 238)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 10
End With

With Me.cmd_Schließen
.BackColor = RGB(255, 255, 255)
.ForeColor = RGB(0, 0, 0)
.Font.Name = "Calibri"
.Font.Size = 10
End With

'_____________Checkbox_____________________________

'With Me.chk_noBud_1
''.Value = Tabelle5.Range("X2").Value
'.BackColor = RGB(255, 255, 255)
'.Font.Name = "Calibri"
'.Font.Size = 10
'.Value = ChrW(&H2713)
'End With

With Me.chk_Mo
'.Value = Tabelle5.Range("Y2").Value
.BackColor = RGB(255, 255, 255)
.Font.Name = "Calibri"
.Font.Size = 10
End With

With Me.chk_Di
'.Value = Tabelle5.Range("Y3").Value
.BackColor = RGB(255, 255, 255)
.Font.Name = "Calibri"
.Font.Size = 10
End With

With Me.chk_Mi
'.Value = Tabelle5.Range("Y4").Value
.BackColor = RGB(255, 255, 255)
.Font.Name = "Calibri"
.Font.Size = 10
End With

With Me.chk_Do
'.Value = Tabelle5.Range("Y5").Value
.BackColor = RGB(255, 255, 255)
.Font.Name = "Calibri"
.Font.Size = 10
End With

With Me.chk_Fr
'.Value = Tabelle5.Range("Y6").Value
.BackColor = RGB(255, 255, 255)
.Font.Name = "Calibri"
.Font.Size = 10
End With

With Me.chk_Sa
'.Value = Tabelle5.Range("Y7").Value
.BackColor = RGB(255, 255, 255)
.Font.Name = "Calibri"
.Font.Size = 10
End With

With Me.chk_So
'.Value = Tabelle5.Range("Y8").Value
.BackColor = RGB(255, 255, 255)
.Font.Name = "Calibri"
.Font.Size = 10
End With

intLGR_Row = Tabelle5.Range("O" & Rows.Count).End(xlUp).row
'****************Filtern************************************

With ActiveWorkbook.Sheets("mapping")
.AutoFilter.Sort.SortFields.Clear
.Range("O2:AF" & intLGR_Row).Sort key1:=.Range("P2:P" & intLGR_Row), order1:=xlAscending, Header:=xlNo
.Range("O2:AF" & intLGR_Row).Sort key1:=.Range("O2:O" & intLGR_Row), order1:=xlAscending, Header:=xlNo
End With

'_____________Listbox_____________________________


myarray = Tabelle5.Range("O2:AF" & intLGR_Row)

With Me.ListBox1
    .Clear
    .ColumnCount = 18
    .ColumnHeads = False
    .ColumnWidths = "70 Pt;80 Pt;56 Pt;72 Pt;55 Pt;36 Pt;51 Pt;72 Pt;72 Pt;77 Pt;75 Pt;36 Pt;36 Pt;36 Pt;36 Pt;36 Pt;36 Pt;36 Pt"
    .List = myarray
    .MultiSelect = 0
End With

End Sub

Private Sub cmd_liste_Click___()

Dim intLGR_Row As Integer

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual

' ***************Datumprüfung***************

Me.txt_saison_von.BackColor = rgbWhite
Me.txt_saison_bis.BackColor = rgbWhite

If Me.txt_saison_von.Value = "" And Me.txt_saison_bis.Value = "" Then
    Me.txt_saison_von.BackColor = rgbPink
    Me.txt_saison_bis.BackColor = rgbPink
    Exit Sub
End If

If Me.txt_saison_von.Value = "" Then
    Me.txt_saison_von.BackColor = rgbPink
    Exit Sub
End If

If Me.txt_saison_bis.Value = "" Then
    Me.txt_saison_bis.BackColor = rgbPink
    Exit Sub
End If

With Me.txt_saison_von
    If CStr(Len(Me.txt_saison_von)) = "6" Then
        Tag = Left(Me.txt_saison_von.Value, 2)
        Monat = Mid(Me.txt_saison_von.Value, 3, 2)
        Jahr = Right(Me.txt_saison_von, 2)
        Me.txt_saison_von = DateSerial(CInt(Jahr), CInt(Monat), CInt(Tag))
        .SetFocus
    ElseIf CStr(Len(Me.txt_saison_von)) = "8" Then
        Tag = Left(Me.txt_saison_von.Value, 2)
        Monat = Mid(Me.txt_saison_von.Value, 3, 2)
        Jahr = Right(Me.txt_saison_von, 4)
        Me.txt_saison_von = DateSerial(CInt(Jahr), CInt(Monat), CInt(Tag))
        .SetFocus
    ElseIf Not IsDate(.Text) Then
        MsgBox "Please enter a <Season start> date. (Dateformat: DDMMYY, DDMMYYYY, DD.MM.YYYY, oder DD/MM/YYYY)"
        '.Value = Date + 30
        .SetFocus
        Exit Sub
    End If
End With

With Me.txt_saison_bis
    If CStr(Len(Me.txt_saison_bis)) = "6" Then
        Tag = Left(Me.txt_saison_bis.Value, 2)
        Monat = Mid(Me.txt_saison_bis.Value, 3, 2)
        Jahr = Right(Me.txt_saison_bis, 2)
        Me.txt_saison_bis = DateSerial(CInt(Jahr), CInt(Monat), CInt(Tag))
        .SetFocus
    ElseIf CStr(Len(Me.txt_saison_bis)) = "8" Then
        Tag = Left(Me.txt_saison_bis.Value, 2)
        Monat = Mid(Me.txt_saison_bis.Value, 3, 2)
        Jahr = Right(Me.txt_saison_bis, 4)
        Me.txt_saison_bis = DateSerial(CInt(Jahr), CInt(Monat), CInt(Tag))
        .SetFocus
    ElseIf Not IsDate(.Text) Then
        MsgBox "Please enter a <Season end> date. (Dateformat: DDMMYY, DDMMYYYY, DD.MM.YYYY, oder DD/MM/YYYY)"
        '.Value = Date + 31
        .SetFocus
        Exit Sub
    End If
End With

If CDate(Me.txt_saison_von.Value) < DateSerial(2021, 1, 1) Or CDate(Me.txt_saison_von.Value) > DateSerial(2024, 12, 31) Then
    MsgBox "The <Season start> date must be between 01.01.2021 and 31.12.2024"
    Me.txt_saison_von.SetFocus
    Exit Sub
End If

If CDate(Me.txt_saison_bis.Value) < DateSerial(2021, 1, 2) Or CDate(Me.txt_saison_bis.Value) > DateSerial(2025, 1, 1) Then
    MsgBox "The <Season end> date must be between 02.01.2021 and 01.01.2025"
    Me.txt_saison_von.SetFocus
    Exit Sub
End If

If Not CDate(Me.txt_saison_von.Value) < CDate(Me.txt_saison_bis.Value) Then
    MsgBox "<Season start> date must be earlier than <Season end>."
    Me.txt_saison_von.Value = ""
    Me.txt_saison_bis.Value = ""
    Exit Sub
End If

intLGR_Row = Range("O" & Rows.Count).End(xlUp).row + 1

ActiveWorkbook.Worksheets("mapping").AutoFilter.Sort.SortFields.Clear
Tabelle5.Range("O2:AF" & intLGR_Row).Sort key1:=Range("P2:P" & intLGR_Row), order1:=xlAscending, Header:=xlNo
Tabelle5.Range("O2:AF" & intLGR_Row).Sort key1:=Range("O2:O" & intLGR_Row), order1:=xlAscending, Header:=xlNo

'Tabelle5.Range("O1:AF" & intLGR_Row).Replace(True, "x", lookat:=xlFormula, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False)

'MyArray = Tabelle5.Range("O1:AF" & intLGR_Row) '.Replace(True, "x", lookat:=xlWhole, SearchOrder:=xlByRows, MatchCase:=True)
'MyArray = Tabelle5.Range("O1:AF" & intLGR_Row).Replace(False, "-", lookat:=xlWhole, SearchOrder:=xlByRows, MatchCase:=True)

'With Me.ListBox1
'    .Clear
'    .ColumnCount = 18
'    .ColumnHeads = False
'    .List = MyArray
'    .MultiSelect = fmMultiSelectExtended
'End With

intLGR_Row = Range("O" & Rows.Count).End(xlUp).row + 1
myarray = Tabelle5.Range("O2:AF" & intLGR_Row)

With Me.ListBox1
    .Clear
    .ColumnCount = 18
    .ColumnHeads = False
    .ColumnWidths = "70 Pt;80 Pt;56 Pt;72 Pt;55 Pt;36 Pt;51 Pt;72 Pt;72 Pt;77 Pt;75 Pt;36 Pt;36 Pt;36 Pt;36 Pt;36 Pt;36 Pt;36 Pt"
    .List = myarray
    .MultiSelect = 0
End With

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic

'MyArray = Tabelle5.Range("O2:AF" & intLGR_Row).Replace("x", True, lookat:=xlWhole, SearchOrder:=xlByRows, MatchCase:=True)
'MyArray = Tabelle5.Range("O2:AF" & intLGR_Row).Replace("-", False, lookat:=xlWhole, SearchOrder:=xlByRows, MatchCase:=True)

End Sub

Public Function OVERLAPS(start1, finish1, start2, finish2) As Variant
 
  If IsDate(start1) And IsDate(finish1) And _
     IsDate(start2) And IsDate(finish2) Then
        OVERLAPS = _
        (start2 >= start1 And start2 <= finish1) Or _
        (finish2 >= start1 And finish2 <= finish1)
  Else
        OVERLAPS = False
  End If

End Function
'
'Public Function OVERLAPS2(Mo1, Mo2, Mo1, Mo2, Mo1, Mo2, Mo1, Do2, Fr1, Fr2, Sa1, Sa2, So1, So2) As Variant
'
'  If IsDate(start1) And IsDate(finish1) And _
'     IsDate(start2) And IsDate(finish2) Then
'        OVERLAPS = _
'        (start2 >= start1 And start2 <= finish1) Or _
'        (finish2 >= start1 And finish2 <= finish1)
'  Else
'        OVERLAPS = False
'  End If
'
'End Function

Public Function Kreuz(strkreuz As String) As Boolean

If strkreuz = ChrW(&H2717) Then
Kreuz = False
Else: Kreuz = True
End If

End Function
