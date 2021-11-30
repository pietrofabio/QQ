VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BGR_Config_v3 
   Caption         =   "BGR Config"
   ClientHeight    =   7308
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   23388
   OleObjectBlob   =   "BGR_Config_v3.frx":0000
End
Attribute VB_Name = "BGR_Config_v3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_liste_Click()

Dim rowVon As Integer
Dim rowBis As Integer
Dim length As Integer

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual

' ***************Datumprüfung***************

Me.txt_saison_von.BackColor = rgbWhite
Me.txt_saison_bis.BackColor = rgbWhite
'Me.txt_Kat_BRO.BackColor = rgbWhite

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
    ElseIf CStr(Len(Me.txt_saison_bis)) = "8" And Mid(Me.txt_saison_bis.Value, 3, 1) = "." Then
        Tag = Left(Me.txt_saison_bis.Value, 2)
        Monat = Mid(Me.txt_saison_bis.Value, 4, 2)
        Jahr = Right(Me.txt_saison_bis, 2)
        Me.txt_saison_bis = DateSerial(CInt(Jahr), CInt(Monat), CInt(Tag))
        .SetFocus
    ElseIf CStr(Len(Me.txt_saison_bis)) = "8" And Mid(Me.txt_saison_bis.Value, 3, 1) <> "." Then
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
    Me.txt_saison_bis.SetFocus
    Exit Sub
End If

If Not CDate(Me.txt_saison_von.Value) < CDate(Me.txt_saison_bis.Value) Then
    MsgBox "<Season start> date must be earlier than <Season end>."
    Me.txt_saison_von.Value = ""
    Me.txt_saison_bis.Value = ""
    Exit Sub
End If

'Blankcell finden
'For Each cell In Tabelle5.Columns("AH:AH").Cells
    'If IsEmpty(cell) = True Then row = cell.row: Exit For
'Next cell

' ***************MainThing***************

Tabelle5.Range("AY2").Value = Me.chk_Mo
Tabelle5.Range("AZ2").Value = Me.chk_Di
Tabelle5.Range("BA2").Value = Me.chk_Mi
Tabelle5.Range("BB2").Value = Me.chk_Do
Tabelle5.Range("BC2").Value = Me.chk_Fr
Tabelle5.Range("BD2").Value = Me.chk_Sa
Tabelle5.Range("BE2").Value = Me.chk_So

rowVon = Tabelle5.Range("AI:AI").Find(what:=CDate(Me.txt_saison_von.Value), LookIn:=xlValues).row
rowBis = Tabelle5.Range("AI:AI").Find(what:=CDate(Me.txt_saison_bis.Value), LookIn:=xlValues).row
length = rowBis - rowVon + 2
Tabelle5.Range("BF2:BU1500").ClearContents

For i = 0 To length - 2  'Schleife wenn richtig
    If WorksheetFunction.HLookup(Tabelle5.Range("AJ" & rowVon + i), Tabelle5.Range("AY1:BE2"), 2, False) = True Then
        Set c = Tabelle5.Range("AJ" & rowVon + i)
        nextrow = Tabelle5.Range("BH1048576").End(xlUp).row + 1
        Tabelle5.Range("BF" & nextrow & ":BU" & nextrow) = _
        Tabelle5.Range("AI" & c.row & ":AX" & c.row).Value
    End If
Next i

On Error Resume Next
myarray = Tabelle5.Range("BF2:BU" & nextrow)

'Listbox
With Me.ListBox1
    .Clear
    .ColumnCount = 16
    .ColumnHeads = False
    .ColumnWidths = "80 Pt;60 Pt;70 Pt;70 Pt;75 Pt;75 Pt;70 Pt;70 Pt;75 Pt;70 Pt;75 Pt;70 Pt;70 Pt;75 Pt;75 Pt;75 Pt;75 Pt;75 Pt;75 Pt"
    .List = myarray
    .MultiSelect = 0
End With

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic

End Sub

Private Sub cmd_UpdateBRO_Click()

Dim i As Integer, j As Integer, rowDow As Integer
Dim row As Integer
Dim rngAnreise As Range, rngAbreise As Range

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual

'Datenprüfung

If Me.txt_BRO.Value = "" And Me.txt_Max_BRO.Value = "" And Me.txt_Fix_BRO.Value = "" _
And Me.txt_Bel_BRO.Value = "" And Me.txt_SpielR_BRO.Value = "" And Me.txt_Min_Rate_BRO.Value = "" Then
    Me.txt_BRO.BackColor = rgbPink
    Me.txt_Max_BRO.BackColor = rgbPink
    Me.txt_Fix_BRO.BackColor = rgbPink
    Me.txt_Bel_BRO.BackColor = rgbPink
    Me.txt_SpielR_BRO.BackColor = rgbPink
    Me.txt_Min_Rate_BRO.BackColor = rgbPink
    Exit Sub
End If

'Discount

Me.txt_BRO.BackColor = rgbWhite

If Me.txt_BRO.Value = "" Then
    GoTo Max_Zi_Check
ElseIf Not IsNumeric(Me.txt_BRO.Value) Then
    Me.txt_BRO.BackColor = rgbPink
    MsgBox "The BRO Discount number must be entered in numeric format."
    Exit Sub
ElseIf Left(CStr(Me.txt_BRO.Value), 2) <> "0," And Left(CStr(Me.txt_BRO.Value), 1) <> "0" Then
    Me.txt_BRO.BackColor = rgbPink
    MsgBox "Der BRO Discount Zahl muss als Dezimalwert eingegeben werden."
    Exit Sub
ElseIf Me.txt_BRO.Value > 0.5 Or Me.txt_BRO.Value < 0 Then
    Me.txt_BRO.BackColor = rgbPink
    MsgBox "Der BRO Discount Zahl ist zu hoch, bzw. zu niedrig."
    Exit Sub
ElseIf Len(Me.txt_BRO.Value) > 4 Then
    Me.txt_BRO.Value = Round(Me.txt_BRO.Value, 2)
End If

Max_Zi_Check:
    
'Max Zi
    
Me.txt_Max_BRO.BackColor = rgbWhite

If Me.txt_Max_BRO.Value = "" Then
    GoTo Fix_Rate_Check
ElseIf Not IsNumeric(Me.txt_Max_BRO.Value) Then
    Me.txt_Max_BRO.BackColor = rgbPink
    MsgBox "Der BRO Max # Zahl muss im Zahlformat eingegeben werden."
    Exit Sub
ElseIf Me.txt_Max_BRO.Value > 200 Or Me.txt_Max_BRO.Value < 0 Then
    Me.txt_Max_BRO.BackColor = rgbPink
    MsgBox "Der BRO Max # Zahl ist zu hoch, bzw. zu niedrig."
    Exit Sub
ElseIf Len(Me.txt_Max_BRO.Value) > 0 Then
    Me.txt_Max_BRO.Value = Round(Me.txt_Max_BRO.Value, 0)
End If

Fix_Rate_Check:
    
'Fix Rate
    
Me.txt_Fix_BRO.BackColor = rgbWhite

If Me.txt_Fix_BRO.Value = "" Then
    GoTo Min_Rate_Check
ElseIf Not IsNumeric(Me.txt_Fix_BRO.Value) Then
    Me.txt_Fix_BRO.BackColor = rgbPink
    MsgBox "Der BRO Fix Rate Zahl muss im Zahlformat eingegeben werden."
    Exit Sub
ElseIf Me.txt_Fix_BRO.Value > 1000 Or Me.txt_Fix_BRO.Value < 0 Then
    Me.txt_Fix_BRO.BackColor = rgbPink
    MsgBox "Der BRO Fix Rate Zahl ist zu hoch, bzw. zu niedrig."
    Exit Sub
ElseIf Len(Me.txt_Fix_BRO.Value) > 0 Then
    Me.txt_Fix_BRO.Value = Round(Me.txt_Fix_BRO.Value, 0)
End If

Min_Rate_Check:

'Min. Rate BRO

Me.txt_Min_Rate_BRO.BackColor = rgbWhite

If Me.txt_Min_Rate_BRO.Value = "" Then
    GoTo Total_Occ_Check
ElseIf Not IsNumeric(Me.txt_Min_Rate_BRO.Value) Then
    Me.txt_Min_Rate_BRO.BackColor = rgbPink
    MsgBox "Der BRO Min. Rate Zahl muss im Zahlformat eingegeben werden."
    Exit Sub
ElseIf Me.txt_Min_Rate_BRO.Value > 1000 Or Me.txt_Min_Rate_BRO.Value < 0 Then
    Me.txt_Min_Rate_BRO.BackColor = rgbPink
    MsgBox "Der BRO Min. Rate Zahl ist zu hoch, bzw. zu niedrig."
    Exit Sub
ElseIf Len(Me.txt_Min_Rate_BRO.Value) > 0 Then
    Me.txt_Min_Rate_BRO.Value = Round(Me.txt_Min_Rate_BRO.Value, 0)
End If

Spielraum_Check:

'Spielraum BRO

Me.txt_SpielR_BRO.BackColor = rgbWhite

If Me.txt_SpielR_BRO.Value = "" Then
    GoTo Total_Occ_Check
ElseIf Not IsNumeric(Me.txt_SpielR_BRO.Value) Then
    Me.txt_SpielR_BRO.BackColor = rgbPink
    MsgBox "Der BRO Spielraumszahl muss im Zahlformat eingegeben werden."
    Exit Sub
ElseIf Me.txt_SpielR_BRO.Value > 100 Or Me.txt_SpielR_BRO.Value < 0 Then
    Me.txt_SpielR_BRO.BackColor = rgbPink
    MsgBox "Der BRO Spielraumszahl ist zu hoch, bzw. zu niedrig."
    Exit Sub
ElseIf Len(Me.txt_SpielR_BRO.Value) > 0 Then
    Me.txt_SpielR_BRO.Value = Round(Me.txt_SpielR_BRO.Value, 0)
End If

Total_Occ_Check:
    
'Total Occ.

Me.txt_Bel_BRO.BackColor = rgbWhite

If Me.txt_Bel_BRO.Value = "" Then
    GoTo Datumprüfung
ElseIf IsNumeric(Me.txt_Bel_BRO.Value) = False Then
    Me.txt_Bel_BRO.BackColor = rgbPink
    MsgBox "Der BRO Total Occ. Zahl muss im Zahlformat (0 - 100) eingegeben werden."
    Exit Sub
ElseIf Me.txt_Bel_BRO.Value < 1 And Me.txt_Bel_BRO.Value > 0 Then
    Me.txt_Bel_BRO.Value = Me.txt_Bel_BRO.Value * 100
ElseIf Me.txt_Bel_BRO.Value < 100 And Me.txt_Bel_BRO.Value > 0 Then
    Me.txt_Bel_BRO.Value = Me.txt_Bel_BRO.Value
ElseIf Me.txt_Bel_BRO.Value > 100 Or Me.txt_Bel_BRO.Value < 0 Then
    Me.txt_Bel_BRO.BackColor = rgbPink
    MsgBox "Der BRO Total Occ. Zahl ist zu hoch, bzw. zu niedrig."
    Exit Sub
ElseIf Len(Me.txt_Bel_BRO.Value) > 0 Then
    Me.txt_Bel_BRO.Value = Round(Me.txt_Bel_BRO.Value, 2)
End If

Datumprüfung:

'Datumprüfung

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
        MsgBox "Sie müssen ein <Saison von> Datum eingeben. (Datumsformat: TTMMJJ, TTMMJJJJ, TT.MM.JJ, TT.MM.YYYY, oder TT/MM/JJJJ)"
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
    ElseIf CStr(Len(Me.txt_saison_bis)) = "8" And Mid(Me.txt_saison_bis.Value, 3, 1) = "." Then
        Tag = Left(Me.txt_saison_bis.Value, 2)
        Monat = Mid(Me.txt_saison_bis.Value, 4, 2)
        Jahr = Right(Me.txt_saison_bis, 2)
        Me.txt_saison_bis = DateSerial(CInt(Jahr), CInt(Monat), CInt(Tag))
        .SetFocus
    ElseIf CStr(Len(Me.txt_saison_bis)) = "8" And Mid(Me.txt_saison_bis.Value, 3, 1) <> "." Then
        Tag = Left(Me.txt_saison_bis.Value, 2)
        Monat = Mid(Me.txt_saison_bis.Value, 3, 2)
        Jahr = Right(Me.txt_saison_bis, 4)
        Me.txt_saison_bis = DateSerial(CInt(Jahr), CInt(Monat), CInt(Tag))
        .SetFocus
    ElseIf Not IsDate(.Text) Then
        MsgBox "Sie müssen ein <Saison bis> Datum eingeben. (Datumsformat: TTMMJJ, TTMMJJJJ, TT.MM.JJ, TT.MM.YYYY, oder TT/MM/JJJJ)"
        '.Value = Date + 31
        .SetFocus
        Exit Sub
    End If
End With

If CDate(Me.txt_saison_von.Value) < DateSerial(2021, 1, 1) Or CDate(Me.txt_saison_von.Value) > DateSerial(2024, 12, 31) Then
    MsgBox "Das <Saison von> Datum muss zwischen 01.01.2021 und 31.12.2024 liegen"
    Me.txt_saison_von.SetFocus
    Exit Sub
End If

If CDate(Me.txt_saison_bis.Value) < DateSerial(2021, 1, 2) Or CDate(Me.txt_saison_bis.Value) > DateSerial(2025, 1, 1) Then
    MsgBox "Das <Saison bis> Datum muss zwischen 02.01.2021 und 01.01.2025 liegen"
    Me.txt_saison_bis.SetFocus
    Exit Sub
End If

If Not CDate(Me.txt_saison_von.Value) < CDate(Me.txt_saison_bis.Value) Then
    MsgBox "<Saison von> Datum muss später sein, als <Saison bis> Datum."
    Me.txt_saison_von.Value = ""
    Me.txt_saison_bis.Value = ""
    Exit Sub
End If

'Selection Ausschließen

For i = 0 To ListBox1.ListCount - 1
    If ListBox1.Selected(i) Then
        'Run Bearbeiten
        Call cmd_bearbeiten_Click
        Exit Sub
    End If
Next i

'Fix Rate BRO

'If Me.txt_Fix_BRO.Value = "" Then
'    Me.txt_Fix_BRO.Value = 0
'ElseIf IsNumeric(Me.txt_Fix_BRO.Value) = False Or Me.txt_Fix_BRO.Value > 800 Then
'    MsgBox "Sie müssen einen Wert eingeben!"
'    Exit Sub
'End If

'***********************MainThing************************

Tabelle5.Range("AY2").Value = Me.chk_Mo
Tabelle5.Range("AZ2").Value = Me.chk_Di
Tabelle5.Range("BA2").Value = Me.chk_Mi
Tabelle5.Range("BB2").Value = Me.chk_Do
Tabelle5.Range("BC2").Value = Me.chk_Fr
Tabelle5.Range("BD2").Value = Me.chk_Sa
Tabelle5.Range("BE2").Value = Me.chk_So

rowVon = Tabelle5.Range("AI:AI").Find(what:=CDate(Me.txt_saison_von.Value), LookIn:=xlValues).row
rowBis = Tabelle5.Range("AI:AI").Find(what:=CDate(Me.txt_saison_bis.Value), LookIn:=xlValues).row
length = rowBis - rowVon + 2
Tabelle5.Range("BF2:BU1500").ClearContents

'***********************Discounts eingeben***********************

For i = 0 To length - 2  'Schleife wenn richtig
    If WorksheetFunction.HLookup(Tabelle5.Range("AJ" & rowVon + i), Tabelle5.Range("AY1:BE2"), 2, False) = True Then
        Set c = Tabelle5.Range("AJ" & rowVon + i)
        nextrow = Tabelle5.Range("BH1048576").End(xlUp).row + 1
        Tabelle5.Range("BF" & nextrow & ":BU" & nextrow) = _
        Tabelle5.Range("AI" & c.row & ":AX" & c.row).Value
        If Me.txt_BRO.Value <> "" Then
            Tabelle5.Range("BP" & nextrow) = Me.txt_BRO.Value * 1
            Tabelle5.Range("AS" & c.row) = Me.txt_BRO.Value * 1
        End If                                                                                                      'Ändern Spiegeltabelle Disc
        If Me.txt_Max_BRO.Value <> "" Then
            Tabelle5.Range("BQ" & nextrow) = Me.txt_Max_BRO.Value * 1                                                   'Ändern Spiegeltabelle Max Zi
            Tabelle5.Range("AT" & c.row) = Me.txt_Max_BRO.Value * 1                                                     'Ändern in Sourcetabelle Max Zi
        End If
        If Me.txt_Fix_BRO.Value <> "" Then
            Tabelle5.Range("BR" & nextrow) = Me.txt_Fix_BRO.Value * 1
            Tabelle5.Range("AU" & c.row) = Me.txt_Fix_BRO.Value * 1
        End If
        If Me.txt_Bel_BRO <> "" Then
            Tabelle5.Range("BS" & nextrow) = Me.txt_Bel_BRO.Value * 1
            Tabelle5.Range("AV" & c.row) = Me.txt_Bel_BRO.Value * 1
        End If
        If Me.txt_Min_Rate_BRO <> "" Then
            Tabelle5.Range("BT" & nextrow) = Me.txt_Min_Rate_BRO.Value * 1
            Tabelle5.Range("AW" & c.row) = Me.txt_Min_Rate_BRO.Value * 1
        End If
        If Me.txt_SpielR_BRO <> "" Then
            Tabelle5.Range("BU" & nextrow) = Me.txt_SpielR_BRO.Value * 1
            Tabelle5.Range("AX" & c.row) = Me.txt_SpielR_BRO.Value * 1
        End If
    End If
Next i

On Error Resume Next
myarray = Tabelle5.Range("BF2:BU" & nextrow)

'Listbox
With Me.ListBox1
    .Clear
    .ColumnCount = 16
    .ColumnHeads = False
    .ColumnWidths = "80 Pt;60 Pt;70 Pt;70 Pt;75 Pt;75 Pt;70 Pt;70 Pt;75 Pt;70 Pt;75 Pt;70 Pt;70 Pt;75 Pt;75 Pt;75 Pt;75 Pt;75 Pt;75 Pt"
    .List = myarray
    .MultiSelect = 0
End With

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic

End Sub

Private Sub cmd_UpdateBME_Click()

Dim i As Integer, j As Integer, rowDow As Integer
Dim row As Integer
Dim rngAnreise As Range, rngAbreise As Range
Dim strDow As String

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual

'Datenprüfung

If Me.txt_BME.Value = "" And Me.txt_Max_BME.Value = "" And Me.txt_Fix_BME.Value = "" _
And Me.txt_Bel_BME.Value = "" And Me.txt_SpielR_BME.Value = "" And Me.txt_Min_Rate_BME.Value = "" Then
    Me.txt_BME.BackColor = rgbPink
    Me.txt_Max_BME.BackColor = rgbPink
    Me.txt_Fix_BME.BackColor = rgbPink
    Me.txt_Bel_BME.BackColor = rgbPink
    Me.txt_SpielR_BME.BackColor = rgbPink
    Me.txt_Min_Rate_BME.BackColor = rgbPink
    Exit Sub
End If

    'Discount
Me.txt_BME.BackColor = rgbWhite

If Me.txt_BME.Value = "" Then
    'Me.txt_BME.Value = 0
    GoTo Max_Zi_Check
ElseIf Not IsNumeric(Me.txt_BME.Value) Then
    Me.txt_BME.BackColor = rgbPink
    MsgBox "Der BME Discount Zahl muss im Zahlformat eingegeben werden."
    Exit Sub
ElseIf Left(CStr(Me.txt_BME.Value), 2) <> "0," And Left(CStr(Me.txt_BME.Value), 1) <> "0" Then
    Me.txt_BME.BackColor = rgbPink
    MsgBox "Der BME Discount Zahl muss als Dezimalwert eingegeben werden."
    Exit Sub
ElseIf Me.txt_BME.Value > 0.5 Or Me.txt_BME.Value < 0 Then
    Me.txt_BME.BackColor = rgbPink
    MsgBox "Der BME Discount Zahl ist zu hoch, bzw. zu niedrig."
    Exit Sub
ElseIf Len(Me.txt_BME.Value) > 4 Then
    Me.txt_BME.Value = Round(Me.txt_BME.Value, 2)
End If

Max_Zi_Check:
    
'Max Zi
    
Me.txt_Max_BME.BackColor = rgbWhite

If Me.txt_Max_BME.Value = "" Then
    'Me.txt_Max_BME.Value = 0
    GoTo Fix_Rate_Check
ElseIf Not IsNumeric(Me.txt_Max_BME.Value) Then
    Me.txt_Max_BME.BackColor = rgbPink
    MsgBox "Der BME Max # Zahl muss im Zahlformat eingegeben werden."
    Exit Sub
ElseIf Me.txt_Max_BME.Value > 200 Or Me.txt_Max_BME.Value < 0 Then
    Me.txt_Max_BME.BackColor = rgbPink
    MsgBox "Der BME Max # Zahl ist zu hoch, bzw. zu niedrig."
    Exit Sub
ElseIf Len(Me.txt_Max_BME.Value) > 0 Then
    Me.txt_Max_BME.Value = Round(Me.txt_Max_BME.Value, 0)
End If

Fix_Rate_Check:
    
'Fix Rate
    
Me.txt_Fix_BME.BackColor = rgbWhite

If Me.txt_Fix_BME.Value = "" Then
    'Me.txt_Fix_BME.Value = 0
    GoTo Total_Occ_Check
ElseIf Not IsNumeric(Me.txt_Fix_BME.Value) Then
    Me.txt_Fix_BME.BackColor = rgbPink
    MsgBox "Der BME Fix Rate Zahl muss im Zahlformat eingegeben werden."
    Exit Sub
ElseIf Me.txt_Fix_BME.Value > 1000 Or Me.txt_Fix_BME.Value < 0 Then
    Me.txt_Fix_BME.BackColor = rgbPink
    MsgBox "Der BME Fix Rate Zahl ist zu hoch, bzw. zu niedrig."
    Exit Sub
ElseIf Len(Me.txt_Fix_BME.Value) > 0 Then
    Me.txt_Fix_BME.Value = Round(Me.txt_Fix_BME.Value, 0)
End If

Min_Rate_Check:

'Min. Rate BME

Me.txt_Min_Rate_BME.BackColor = rgbWhite

If Me.txt_Min_Rate_BME.Value = "" Then
    GoTo Spielraum_Check
ElseIf Not IsNumeric(Me.txt_Min_Rate_BME.Value) Then
    Me.txt_Min_Rate_BME.BackColor = rgbPink
    MsgBox "Der BME Min. Rate Zahl muss im Zahlformat eingegeben werden."
    Exit Sub
ElseIf Me.txt_Min_Rate_BME.Value > 1000 Or Me.txt_Min_Rate_BME.Value < 0 Then
    Me.txt_Min_Rate_BME.BackColor = rgbPink
    MsgBox "Der BME Min. Rate Zahl ist zu hoch, bzw. zu niedrig."
    Exit Sub
ElseIf Len(Me.txt_Min_Rate_BME.Value) > 0 Then
    Me.txt_Min_Rate_BME.Value = Round(Me.txt_Min_Rate_BME.Value, 0)
End If

Spielraum_Check:

'Spielraum BME

Me.txt_SpielR_BME.BackColor = rgbWhite

If Me.txt_SpielR_BME.Value = "" Then
    GoTo Total_Occ_Check
ElseIf Not IsNumeric(Me.txt_SpielR_BME.Value) Then
    Me.txt_SpielR_BME.BackColor = rgbPink
    MsgBox "Der BME Spielraumszahl muss im Zahlformat eingegeben werden."
    Exit Sub
ElseIf Me.txt_SpielR_BME.Value > 100 Or Me.txt_SpielR_BME.Value < 0 Then
    Me.txt_SpielR_BME.BackColor = rgbPink
    MsgBox "Der BME Spielraumszahl ist zu hoch, bzw. zu niedrig."
    Exit Sub
ElseIf Len(Me.txt_SpielR_BME.Value) > 0 Then
    Me.txt_SpielR_BME.Value = Round(Me.txt_SpielR_BME.Value, 0)
End If

Total_Occ_Check:
    
'Total Occ.
    
Me.txt_Bel_BME.BackColor = rgbWhite

If Me.txt_Bel_BME.Value = "" Then
    'Me.txt_Bel_BME.Value = 0
    GoTo Datumprüfung
ElseIf IsNumeric(Me.txt_Bel_BME.Value) = False Then
    Me.txt_Bel_BME.BackColor = rgbPink
    MsgBox "Der BME Total Occ. Zahl muss im Zahlformat eingegeben werden."
    Exit Sub
ElseIf Me.txt_Bel_BME.Value < 1 And Me.txt_Bel_BME.Value > 0 Then
    Me.txt_Bel_BME.Value = Me.txt_Bel_BME.Value * 100
ElseIf Me.txt_Bel_BME.Value < 100 And Me.txt_Bel_BME.Value > 0 Then
    Me.txt_Bel_BME.Value = Me.txt_Bel_BME.Value
ElseIf Me.txt_Bel_BME.Value > 100 Or Me.txt_Bel_BME.Value < 0 Then
    Me.txt_Bel_BME.BackColor = rgbPink
    MsgBox "Der BME Total Occ. Zahl ist zu hoch, bzw. zu niedrig."
    Exit Sub
ElseIf Len(Me.txt_Bel_BME.Value) > 0 Then
    Me.txt_Bel_BME.Value = Round(Me.txt_Bel_BME.Value, 2)
End If

Datumprüfung:

'Datumprüfung

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
        MsgBox "Sie müssen ein <Saison von> Datum eingeben. (Datumsformat: TTMMJJ, TTMMJJJJ, TT.MM.JJ, TT.MM.YYYY, oder TT/MM/JJJJ)"
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
    ElseIf CStr(Len(Me.txt_saison_bis)) = "8" And Mid(Me.txt_saison_bis.Value, 3, 1) = "." Then
        Tag = Left(Me.txt_saison_bis.Value, 2)
        Monat = Mid(Me.txt_saison_bis.Value, 4, 2)
        Jahr = Right(Me.txt_saison_bis, 2)
        Me.txt_saison_bis = DateSerial(CInt(Jahr), CInt(Monat), CInt(Tag))
        .SetFocus
    ElseIf CStr(Len(Me.txt_saison_bis)) = "8" And Mid(Me.txt_saison_bis.Value, 3, 1) <> "." Then
        Tag = Left(Me.txt_saison_bis.Value, 2)
        Monat = Mid(Me.txt_saison_bis.Value, 3, 2)
        Jahr = Right(Me.txt_saison_bis, 4)
        Me.txt_saison_bis = DateSerial(CInt(Jahr), CInt(Monat), CInt(Tag))
        .SetFocus
    ElseIf Not IsDate(.Text) Then
        MsgBox "Sie müssen ein <Saison bis> Datum eingeben. (Datumsformat: TTMMJJ, TTMMJJJJ, TT.MM.JJ, TT.MM.YYYY, oder TT/MM/JJJJ)"
        '.Value = Date + 31
        .SetFocus
        Exit Sub
    End If
End With

If CDate(Me.txt_saison_von.Value) < DateSerial(2021, 1, 1) Or CDate(Me.txt_saison_von.Value) > DateSerial(2024, 12, 31) Then
    MsgBox "Das <Saison von> Datum muss zwischen 01.01.2021 und 31.12.2024 liegen"
    Me.txt_saison_von.SetFocus
    Exit Sub
End If

If CDate(Me.txt_saison_bis.Value) < DateSerial(2021, 1, 2) Or CDate(Me.txt_saison_bis.Value) > DateSerial(2025, 1, 1) Then
    MsgBox "Das <Saison bis> Datum muss zwischen 02.01.2021 und 01.01.2025 liegen"
    Me.txt_saison_bis.SetFocus
    Exit Sub
End If

If Not CDate(Me.txt_saison_von.Value) < CDate(Me.txt_saison_bis.Value) Then
    MsgBox "<Saison von> Datum muss später sein, als <Saison bis> Datum."
    Me.txt_saison_von.Value = ""
    Me.txt_saison_bis.Value = ""
    Exit Sub
End If

'Selection Ausschließen

For i = 0 To ListBox1.ListCount - 1
    If ListBox1.Selected(i) Then
        'Run Bearbeiten
        Call cmd_bearbeiten_Click
        Exit Sub
    End If
Next i

''Fix Rate BME
'
'If Me.txt_Fix_BME.Value = "" Then
''    Me.txt_Fix_BME.Value = 0
'ElseIf IsNumeric(Me.txt_Fix_BME.Value) = False Or Me.txt_Fix_BME.Value > 800 Then
'    MsgBox "Sie müssen einen Wert eingeben!"
'    Exit Sub
'End If

'***********************MainThing************************

Tabelle5.Range("AY2").Value = Me.chk_Mo
Tabelle5.Range("AZ2").Value = Me.chk_Di
Tabelle5.Range("BA2").Value = Me.chk_Mi
Tabelle5.Range("BB2").Value = Me.chk_Do
Tabelle5.Range("BC2").Value = Me.chk_Fr
Tabelle5.Range("BD2").Value = Me.chk_Sa
Tabelle5.Range("BE2").Value = Me.chk_So

rowVon = Tabelle5.Range("AI:AI").Find(what:=CDate(Me.txt_saison_von.Value), LookIn:=xlValues).row
rowBis = Tabelle5.Range("AI:AI").Find(what:=CDate(Me.txt_saison_bis.Value), LookIn:=xlValues).row
length = rowBis - rowVon + 2
Tabelle5.Range("BF2:BU1500").ClearContents

'***********************Discounts eingeben***********************

For i = 0 To length - 2  'Schleife wenn richtig
    If WorksheetFunction.HLookup(Tabelle5.Range("AJ" & rowVon + i), Tabelle5.Range("AY1:BE2"), 2, False) = True Then
        Set c = Tabelle5.Range("AJ" & rowVon + i)
        nextrow = Tabelle5.Range("BH1048576").End(xlUp).row + 1
        Tabelle5.Range("BF" & nextrow & ":BU" & nextrow) = _
        Tabelle5.Range("AI" & c.row & ":AX" & c.row).Value
        If Me.txt_BME.Value <> "" Then
            Tabelle5.Range("BH" & nextrow) = Me.txt_BME.Value * 1                                                          'Ändern Spiegeltabelle Disc
            Tabelle5.Range("AK" & c.row) = Me.txt_BME.Value * 1                                                         'Ändern in Sourcetabelle Disc
        End If
        If Me.txt_Max_BME.Value <> "" Then
            Tabelle5.Range("BI" & nextrow) = Me.txt_Max_BME.Value * 1                                                    'Ändern Spiegeltabelle Max Zi
            Tabelle5.Range("AL" & c.row) = Me.txt_Max_BME.Value * 1                                                     'Ändern in Sourcetabelle Max Zi
        End If
        If Me.txt_Fix_BME.Value <> "" Then
            Tabelle5.Range("BJ" & nextrow) = Me.txt_Fix_BME.Value * 1                                                    'Ändern Spiegeltabelle Max Zi
            Tabelle5.Range("AM" & c.row) = Me.txt_Fix_BME.Value * 1                                                     'Ändern in Sourcetabelle Max Zi
        End If
        If Me.txt_Bel_BME.Value <> "" Then
            Tabelle5.Range("BK" & nextrow) = Round(Me.txt_Bel_BME.Value, 2) * 1                                                   'Ändern Spiegeltabelle Max Zi
            Tabelle5.Range("AN" & c.row) = Round(Me.txt_Bel_BME.Value, 2) * 1                                                    'Ändern in Sourcetabelle Max Zi
        End If
        If Me.txt_Min_Rate_BME <> "" Then
            Tabelle5.Range("BL" & nextrow) = Me.txt_Min_Rate_BME.Value * 1
            Tabelle5.Range("AO" & c.row) = Me.txt_Min_Rate_BME.Value * 1
        End If
        If Me.txt_SpielR_BME <> "" Then
            Tabelle5.Range("BM" & nextrow) = Me.txt_SpielR_BME.Value * 1
            Tabelle5.Range("AP" & c.row) = Me.txt_SpielR_BME.Value * 1
        End If
    End If
Next i

On Error Resume Next
myarray = Tabelle5.Range("BF2:BU" & nextrow)

'Listbox
With Me.ListBox1
    .Clear
    .ColumnCount = 16
    .ColumnHeads = False
    .ColumnWidths = "80 Pt;60 Pt;70 Pt;70 Pt;75 Pt;75 Pt;70 Pt;70 Pt;75 Pt;70 Pt;75 Pt;70 Pt;70 Pt;75 Pt;75 Pt;75 Pt;75 Pt;75 Pt;75 Pt"
    .List = myarray
    .MultiSelect = 0
End With

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic

End Sub

Private Sub cmd_Update_BGR_DZ_Click()

Dim i As Integer, j As Integer, rowDow As Integer
Dim row As Integer
Dim rngAnreise As Range, rngAbreise As Range
Dim strDow As String

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual

'Datenprüfung

If Me.txt_bgr_dz.Value = "" Then
    Me.txt_bgr_dz.BackColor = rgbPink
    Exit Sub
End If

Tabelle5.Range("K16").Value = Me.txt_bgr_dz.Value

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic

End Sub

Private Sub cmd_UpdateKat_Click()

Dim i As Integer, j As Integer, rowDow As Integer
Dim row As Integer
Dim rngAnreise As Range, rngAbreise As Range
Dim strDow As String

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual

'Datenprüfung

If Me.txt_Kat.Value = "" Then
    Me.txt_Kat.BackColor = rgbPink
    Exit Sub
End If

'Kategorienaufschlagsprüfung

Me.txt_Kat.BackColor = rgbWhite

If Not IsNumeric(Me.txt_Kat.Value) Then
    Me.txt_Kat.BackColor = rgbPink
    MsgBox "Der Kategorienaufschlag muss im Zahlformat eingegeben werden."
    Exit Sub
ElseIf Me.txt_Kat.Value > 200 Or Me.txt_Kat.Value < 0 Then
    Me.txt_Kat.BackColor = rgbPink
    MsgBox "Der Kategorienaufschlag ist zu hoch, bzw. zu niedrig."
    Exit Sub
Else
    Me.txt_Kat.Value = Round(Me.txt_Kat.Value, 0)
End If

'Datumprüfung

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
        MsgBox "Sie müssen ein <Saison von> Datum eingeben. (Datumsformat: TTMMJJ, TTMMJJJJ, TT.MM.JJ, TT.MM.YYYY, oder TT/MM/JJJJ)"
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
    ElseIf CStr(Len(Me.txt_saison_bis)) = "8" And Mid(Me.txt_saison_bis.Value, 3, 1) = "." Then
        Tag = Left(Me.txt_saison_bis.Value, 2)
        Monat = Mid(Me.txt_saison_bis.Value, 4, 2)
        Jahr = Right(Me.txt_saison_bis, 2)
        Me.txt_saison_bis = DateSerial(CInt(Jahr), CInt(Monat), CInt(Tag))
        .SetFocus
    ElseIf CStr(Len(Me.txt_saison_bis)) = "8" And Mid(Me.txt_saison_bis.Value, 3, 1) <> "." Then
        Tag = Left(Me.txt_saison_bis.Value, 2)
        Monat = Mid(Me.txt_saison_bis.Value, 3, 2)
        Jahr = Right(Me.txt_saison_bis, 4)
        Me.txt_saison_bis = DateSerial(CInt(Jahr), CInt(Monat), CInt(Tag))
        .SetFocus
    ElseIf Not IsDate(.Text) Then
        MsgBox "Sie müssen ein <Saison bis> Datum eingeben. (Datumsformat: TTMMJJ, TTMMJJJJ, TT.MM.JJ, TT.MM.YYYY, oder TT/MM/JJJJ)"
        '.Value = Date + 31
        .SetFocus
        Exit Sub
    End If
End With

If CDate(Me.txt_saison_von.Value) < DateSerial(2021, 1, 1) Or CDate(Me.txt_saison_von.Value) > DateSerial(2024, 12, 31) Then
    MsgBox "Das <Saison von> Datum muss zwischen 01.01.2021 und 31.12.2024 liegen"
    Me.txt_saison_von.SetFocus
    Exit Sub
End If

If CDate(Me.txt_saison_bis.Value) < DateSerial(2021, 1, 2) Or CDate(Me.txt_saison_bis.Value) > DateSerial(2025, 1, 1) Then
    MsgBox "Das <Saison bis> Datum muss zwischen 02.01.2021 und 01.01.2025 liegen"
    Me.txt_saison_bis.SetFocus
    Exit Sub
End If

If Not CDate(Me.txt_saison_von.Value) < CDate(Me.txt_saison_bis.Value) Then
    MsgBox "<Saison von> Datum muss später sein, als <Saison bis> Datum."
    Me.txt_saison_von.Value = ""
    Me.txt_saison_bis.Value = ""
    Exit Sub
End If

'Selection Ausschließen

For i = 0 To ListBox1.ListCount - 1
    If ListBox1.Selected(i) Then
        Call cmd_bearbeiten_Click
        Exit Sub
    End If
Next i

'***********************MainThing************************

Tabelle5.Range("AY2").Value = Me.chk_Mo
Tabelle5.Range("AZ2").Value = Me.chk_Di
Tabelle5.Range("BA2").Value = Me.chk_Mi
Tabelle5.Range("BB2").Value = Me.chk_Do
Tabelle5.Range("BC2").Value = Me.chk_Fr
Tabelle5.Range("BD2").Value = Me.chk_Sa
Tabelle5.Range("BE2").Value = Me.chk_So

rowVon = Tabelle5.Range("AI:AI").Find(what:=CDate(Me.txt_saison_von.Value), LookIn:=xlValues).row
rowBis = Tabelle5.Range("AI:AI").Find(what:=CDate(Me.txt_saison_bis.Value), LookIn:=xlValues).row
length = rowBis - rowVon + 2
Tabelle5.Range("BF2:BU1500").ClearContents

'***********************Discounts eingeben***********************
For i = 0 To length - 2  'Schleife wenn richtig
    If WorksheetFunction.HLookup(Tabelle5.Range("AJ" & rowVon + i), Tabelle5.Range("AY1:BE2"), 2, False) = True Then
        Set c = Tabelle5.Range("AJ" & rowVon + i)
        nextrow = Tabelle5.Range("BH1048576").End(xlUp).row + 1
        Tabelle5.Range("BF" & nextrow & ":BU" & nextrow) = _
        Tabelle5.Range("AI" & c.row & ":AX" & c.row).Value
        If Me.txt_Kat.Value <> "" Then
            Tabelle5.Range("BO" & nextrow) = Me.txt_Kat.Value * 1
            Tabelle5.Range("AR" & c.row) = Me.txt_Kat.Value * 1
        End If
        'If Me.chk_plenum = True Then
        '    Tabelle5.Range("BD" & nextrow) = ChrW(&H2713)
        '    Tabelle5.Range("AM" & c.row) = ChrW(&H2713)
        'ElseIf Me.chk_plenum = False Then
        '    Tabelle5.Range("BD" & nextrow) = ChrW(&H2717)
        '    Tabelle5.Range("AM" & c.row) = ChrW(&H2717)
        'End If
    End If
Next i

On Error Resume Next
myarray = Tabelle5.Range("BF2:BU" & nextrow)

'Listbox
With Me.ListBox1
    .Clear
    .ColumnCount = 16
    .ColumnHeads = False
    .ColumnWidths = "80 Pt;60 Pt;70 Pt;70 Pt;75 Pt;75 Pt;70 Pt;70 Pt;75 Pt;70 Pt;75 Pt;70 Pt;70 Pt;75 Pt;75 Pt;75 Pt;75 Pt;75 Pt;75 Pt"
    .List = myarray
    .MultiSelect = 0
End With

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic

End Sub

Private Sub cmd_UpdatePlenum_Click()

Dim i As Integer, j As Integer, rowDow As Integer
Dim row As Integer
Dim rngAnreise As Range, rngAbreise As Range
Dim strDow As String

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual

'Datumprüfung

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
        MsgBox "Sie müssen ein <Saison von> Datum eingeben. (Datumsformat: TTMMJJ, TTMMJJJJ, TT.MM.JJ, TT.MM.YYYY, oder TT/MM/JJJJ)"
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
    ElseIf CStr(Len(Me.txt_saison_bis)) = "8" And Mid(Me.txt_saison_bis.Value, 3, 1) = "." Then
        Tag = Left(Me.txt_saison_bis.Value, 2)
        Monat = Mid(Me.txt_saison_bis.Value, 4, 2)
        Jahr = Right(Me.txt_saison_bis, 2)
        Me.txt_saison_bis = DateSerial(CInt(Jahr), CInt(Monat), CInt(Tag))
        .SetFocus
    ElseIf CStr(Len(Me.txt_saison_bis)) = "8" And Mid(Me.txt_saison_bis.Value, 3, 1) <> "." Then
        Tag = Left(Me.txt_saison_bis.Value, 2)
        Monat = Mid(Me.txt_saison_bis.Value, 3, 2)
        Jahr = Right(Me.txt_saison_bis, 4)
        Me.txt_saison_bis = DateSerial(CInt(Jahr), CInt(Monat), CInt(Tag))
        .SetFocus
    ElseIf Not IsDate(.Text) Then
        MsgBox "Sie müssen ein <Saison bis> Datum eingeben. (Datumsformat: TTMMJJ, TTMMJJJJ, TT.MM.JJ, TT.MM.YYYY, oder TT/MM/JJJJ)"
        '.Value = Date + 31
        .SetFocus
        Exit Sub
    End If
End With

If CDate(Me.txt_saison_von.Value) < DateSerial(2021, 1, 1) Or CDate(Me.txt_saison_von.Value) > DateSerial(2024, 12, 31) Then
    MsgBox "Das <Saison von> Datum muss zwischen 01.01.2021 und 31.12.2024 liegen"
    Me.txt_saison_von.SetFocus
    Exit Sub
End If

If CDate(Me.txt_saison_bis.Value) < DateSerial(2021, 1, 2) Or CDate(Me.txt_saison_bis.Value) > DateSerial(2025, 1, 1) Then
    MsgBox "Das <Saison bis> Datum muss zwischen 02.01.2021 und 01.01.2025 liegen"
    Me.txt_saison_bis.SetFocus
    Exit Sub
End If

If Not CDate(Me.txt_saison_von.Value) < CDate(Me.txt_saison_bis.Value) Then
    MsgBox "<Saison von> Datum muss später sein, als <Saison bis> Datum."
    Me.txt_saison_von.Value = ""
    Me.txt_saison_bis.Value = ""
    Exit Sub
End If

'Selection Ausschließen

For i = 0 To ListBox1.ListCount - 1
    If ListBox1.Selected(i) Then
        Call cmd_bearbeiten_Click
        Exit Sub
    End If
Next i

'***********************MainThing************************

Tabelle5.Range("AY2").Value = Me.chk_Mo
Tabelle5.Range("AZ2").Value = Me.chk_Di
Tabelle5.Range("BA2").Value = Me.chk_Mi
Tabelle5.Range("BB2").Value = Me.chk_Do
Tabelle5.Range("BC2").Value = Me.chk_Fr
Tabelle5.Range("BD2").Value = Me.chk_Sa
Tabelle5.Range("BE2").Value = Me.chk_So

rowVon = Tabelle5.Range("AI:AI").Find(what:=CDate(Me.txt_saison_von.Value), LookIn:=xlValues).row
rowBis = Tabelle5.Range("AI:AI").Find(what:=CDate(Me.txt_saison_bis.Value), LookIn:=xlValues).row
length = rowBis - rowVon + 2
Tabelle5.Range("BF2:BU1500").ClearContents

'***********************Plenum Protects eingeben***********************

For i = 0 To length - 2  'Schleife wenn richtig
    If WorksheetFunction.HLookup(Tabelle5.Range("AJ" & rowVon + i), Tabelle5.Range("AY1:BE2"), 2, False) = True Then
        Set c = Tabelle5.Range("AJ" & rowVon + i)
        nextrow = Tabelle5.Range("BH1048576").End(xlUp).row + 1
        Tabelle5.Range("BF" & nextrow & ":BU" & nextrow) = _
        Tabelle5.Range("AI" & c.row & ":AX" & c.row).Value
        If Me.chk_plenum = True Then
            Tabelle5.Range("BN" & nextrow) = ChrW(&H2713)
            Tabelle5.Range("AQ" & c.row) = ChrW(&H2713)
        ElseIf Me.chk_plenum = False Then
            Tabelle5.Range("BN" & nextrow) = ChrW(&H2717)
            Tabelle5.Range("AQ" & c.row) = ChrW(&H2717)
        End If
    End If
Next i

On Error Resume Next
myarray = Tabelle5.Range("BF2:BU" & nextrow)

'Listbox
With Me.ListBox1
    .Clear
    .ColumnCount = 16
    .ColumnHeads = False
    .ColumnWidths = "80 Pt;60 Pt;70 Pt;70 Pt;75 Pt;75 Pt;70 Pt;70 Pt;75 Pt;70 Pt;75 Pt;70 Pt;70 Pt;75 Pt;75 Pt;75 Pt;75 Pt;75 Pt;75 Pt"
    .List = myarray
    .MultiSelect = 0
End With

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic

End Sub

Private Sub cmd_bearbeiten_Click()

Dim intBGR_Row As Integer
Dim intBearbeiten_Row As Integer
Dim strkreuz As String
Dim rowVon As Long, rowBis As Long, length As Long

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual

'rowBis = Tabelle5.Range("AH:AH").Find(What:=CDate(Me.txt_saison_bis.Value), LookIn:=xlValues).row

'Debug.Print ListBox1.List(i)

For i = 0 To ListBox1.ListCount - 1
    If ListBox1.Selected(i) Then
        intBearbeiten_Row = Tabelle5.Range("AI:AI").Find(what:=CDate(ListBox1.List(i)), LookIn:=xlValues).row: Exit For
        'intBearbeiten_Row = Tabelle5.Cells(i + rowVon, 35).row: Exit For
'        With lbl_Info
'        .Caption = ""
'        End With: Exit For
    ElseIf i = ListBox1.ListCount - 1 Then
'    With lbl_Info
'    .Caption = "Eine Zeile muss ausgewählt werden."
'    End With
    Exit Sub
    End If
Next i

With Tabelle5
'Me.txt_saison_von.Value = .Cells(intBearbeiten_Row, 34).Value
'Me.txt_saison_bis.Value = .Cells(intBearbeiten_Row, 34).Value
If Me.txt_BME.Value <> "" Then
    .Cells(intBearbeiten_Row, 37).Value = Me.txt_BME.Value * 1
End If
If Me.txt_Max_BME.Value <> "" Then
.Cells(intBearbeiten_Row, 38).Value = Me.txt_Max_BME.Value * 1
End If
If Me.txt_Fix_BME.Value <> "" Then
.Cells(intBearbeiten_Row, 39).Value = Me.txt_Fix_BME.Value * 1
End If
If Me.txt_Bel_BME.Value <> "" Then
.Cells(intBearbeiten_Row, 40).Value = Me.txt_Bel_BME.Value * 1
End If
If Me.txt_Min_Rate_BME.Value <> "" Then
.Cells(intBearbeiten_Row, 41).Value = Me.txt_Min_Rate_BME.Value * 1
End If
If Me.txt_SpielR_BME.Value <> "" Then
.Cells(intBearbeiten_Row, 42).Value = Me.txt_SpielR_BME.Value * 1
End If
If Me.chk_plenum.Value = True Then
.Cells(intBearbeiten_Row, 43).Value = ChrW(&H2713)
Else
.Cells(intBearbeiten_Row, 43).Value = ChrW(&H2717)
End If

'If Me.txt_Kat.Value = "" Then
'    Me.txt_Kat.Value = 0
'End If

If Me.txt_Kat.Value <> "" Then
.Cells(intBearbeiten_Row, 44).Value = Me.txt_Kat.Value * 1
End If
If Me.txt_BRO.Value <> "" Then
.Cells(intBearbeiten_Row, 45).Value = Me.txt_BRO.Value * 1
End If
If Me.txt_Max_BRO.Value <> "" Then
.Cells(intBearbeiten_Row, 46).Value = Me.txt_Max_BRO.Value * 1
End If
If Me.txt_Fix_BRO.Value <> "" Then
.Cells(intBearbeiten_Row, 47).Value = Me.txt_Fix_BRO.Value * 1
End If
If Me.txt_Bel_BRO.Value <> "" Then
.Cells(intBearbeiten_Row, 48).Value = Me.txt_Bel_BRO.Value * 1
End If
If Me.txt_Min_Rate_BRO.Value <> "" Then
.Cells(intBearbeiten_Row, 49).Value = Me.txt_Min_Rate_BRO.Value * 1
End If
If Me.txt_SpielR_BRO.Value <> "" Then
.Cells(intBearbeiten_Row, 50).Value = Me.txt_SpielR_BRO.Value * 1
End If

End With

'***********************MainThing************************

rowVon = Tabelle5.Range("AI:AI").Find(what:=CDate(Me.txt_saison_von.Value), LookIn:=xlValues).row
rowBis = Tabelle5.Range("AI:AI").Find(what:=CDate(Me.txt_saison_bis.Value), LookIn:=xlValues).row
length = rowBis - rowVon + 2
Tabelle5.Range("BF2:BU1500").ClearContents

'***********************Discounts eingeben***********************
For i = 0 To length - 2  'Schleife wenn richtig
    If WorksheetFunction.HLookup(Tabelle5.Range("AJ" & rowVon + i), Tabelle5.Range("AY1:BE2"), 2, False) = True Then
        Set c = Tabelle5.Range("AJ" & rowVon + i)
        nextrow = Tabelle5.Range("BH1048576").End(xlUp).row + 1
        Tabelle5.Range("BF" & nextrow & ":BU" & nextrow) = _
        Tabelle5.Range("AI" & c.row & ":AX" & c.row).Value
'        If Me.chk_plenum = True Then
'            Tabelle5.Range("BD" & nextrow) = ChrW(&H2713)
'            Tabelle5.Range("AM" & c.row) = ChrW(&H2713)
'        ElseIf Me.chk_plenum = False Then
'            Tabelle5.Range("BD" & nextrow) = ChrW(&H2717)
'            Tabelle5.Range("AM" & c.row) = ChrW(&H2717)
'        End If
    End If
Next i

myarray = Tabelle5.Range("BF2:BU" & nextrow)

'Listbox
With Me.ListBox1
    .Clear
    .ColumnCount = 16
    .ColumnHeads = False
    .ColumnWidths = "80 Pt;60 Pt;70 Pt;70 Pt;75 Pt;75 Pt;70 Pt;70 Pt;75 Pt;70 Pt;75 Pt;70 Pt;70 Pt;75 Pt;75 Pt;75 Pt;75 Pt;75 Pt;75 Pt"
    .List = myarray
    .MultiSelect = 0
End With

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic

End Sub

Private Sub ListBox1_Click()

Dim intBGR_Row As Integer
Dim intBearbeiten_Row As Integer
Dim strkreuz As String
Dim rowVon As Long, rowBis As Long, length As Long

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual

Me.txt_saison_von.BackColor = rgbWhite
Me.txt_saison_bis.BackColor = rgbWhite
Me.txt_BME.BackColor = rgbWhite
Me.txt_Max_BME.BackColor = rgbWhite
Me.txt_Fix_BME.BackColor = rgbWhite
Me.txt_BRO.BackColor = rgbWhite
Me.txt_Max_BRO.BackColor = rgbWhite
Me.txt_Fix_BRO.BackColor = rgbWhite
Me.txt_Kat.BackColor = rgbWhite

'intBGR_Row = Tabelle5.Range("O1").End(xlDown).row + 1

'***********************MainThing************************

rowVon = Tabelle5.Range("AI:AI").Find(what:=CDate(Me.txt_saison_von.Value), LookIn:=xlValues).row
'rowBis = Tabelle5.Range("AH:AH").Find(What:=CDate(Me.txt_saison_bis.Value), LookIn:=xlValues).row
'length = rowBis - rowVon + 2
'Tabelle5.Range("AY2:BH800").ClearContents

For i = 0 To ListBox1.ListCount - 1
    If ListBox1.Selected(i) Then
        intBearbeiten_Row = Tabelle5.Range("AI:AI").Find(what:=CDate(ListBox1.List(i)), LookIn:=xlValues).row: Exit For
        'intBearbeiten_Row = Tabelle5.Cells(i + (rowVon), 35).row: Exit For
'        With lbl_Info
'        .Caption = ""
'        End With: Exit For
    ElseIf i = ListBox1.ListCount - 1 Then
'    With lbl_Info
'    .Caption = "Eine Zeile muss ausgewählt werden."
'    End With
    Exit Sub
    End If
Next i

With Tabelle5
'Me.txt_saison_von.Value = .Cells(intBearbeiten_Row, 34).Value
'Me.txt_saison_bis.Value = .Cells(intBearbeiten_Row, 34).Value
Me.txt_BME.Value = .Cells(intBearbeiten_Row, 37).Value
Me.txt_Max_BME.Value = .Cells(intBearbeiten_Row, 38).Value
Me.txt_Fix_BME.Value = .Cells(intBearbeiten_Row, 39).Value
Me.txt_Bel_BME.Value = .Cells(intBearbeiten_Row, 40).Value
Me.txt_Min_Rate_BME.Value = .Cells(intBearbeiten_Row, 41).Value
Me.txt_SpielR_BME.Value = .Cells(intBearbeiten_Row, 42).Value

Me.chk_Mo = False
Me.chk_Di = False
Me.chk_Mi = False
Me.chk_Do = False
Me.chk_Fr = False
Me.chk_Sa = False
Me.chk_So = False

Select Case .Cells(intBearbeiten_Row, 36).Value
    Case "Mon"
    Me.chk_Mo = True
    Case "Tue"
    Me.chk_Di = True
    Case "Wed"
    Me.chk_Mi = True
    Case "Thu"
    Me.chk_Do = True
    Case "Fri"
    Me.chk_Fr = True
    Case "Sat"
    Me.chk_Sa = True
    Case "Sun"
    Me.chk_So = True
End Select

If .Cells(intBearbeiten_Row, 43).Value = ChrW(&H2713) Then
Me.chk_plenum.Value = True
Else
Me.chk_plenum.Value = False
End If

Me.txt_Kat.Value = .Cells(intBearbeiten_Row, 44).Value
Me.txt_BRO.Value = .Cells(intBearbeiten_Row, 45).Value
Me.txt_Max_BRO.Value = .Cells(intBearbeiten_Row, 46).Value
Me.txt_Fix_BRO.Value = .Cells(intBearbeiten_Row, 47).Value
Me.txt_Bel_BRO.Value = .Cells(intBearbeiten_Row, 48).Value
Me.txt_Min_Rate_BRO.Value = .Cells(intBearbeiten_Row, 49).Value
Me.txt_SpielR_BRO.Value = .Cells(intBearbeiten_Row, 50).Value

'strkreuz = .Cells(intBearbeiten_Row, 22).Value
'Me.chk_noBud_1.Value = Kreuz(strkreuz)
'strkreuz = .Cells(intBearbeiten_Row, 23).Value
'Me.chk_Mo.Value = Kreuz(strkreuz)
'strkreuz = .Cells(intBearbeiten_Row, 24).Value
'Me.chk_Di.Value = Kreuz(strkreuz)
'strkreuz = .Cells(intBearbeiten_Row, 25).Value
'Me.chk_Mi.Value = Kreuz(strkreuz)
'strkreuz = .Cells(intBearbeiten_Row, 26).Value
'Me.chk_Do.Value = Kreuz(strkreuz)
'strkreuz = .Cells(intBearbeiten_Row, 27).Value
'Me.chk_Fr.Value = Kreuz(strkreuz)
'strkreuz = .Cells(intBearbeiten_Row, 28).Value
'Me.chk_Sa.Value = Kreuz(strkreuz)
'strkreuz = .Cells(intBearbeiten_Row, 29).Value
'Me.chk_So.Value = Kreuz(strkreuz)

End With

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic

End Sub

Private Sub cmd_Schließen_Click()
Unload BGR_Config_v3
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

Private Sub UserForm_Initialize()

Dim lastrow As Integer

'___________________________General_________________________


'Me.Caption = "Quick Quote von " & Environ("UserName") & " vom " & Date
'Me.BackColor = RGB(240, 240, 240)
'Me.StartUpPosition = 3


'___________________________CommandButton_________________________

'.........................Page1...........................

With Me.txt_saison_von
.Value = "01.01.2021"
End With

With Me.txt_saison_bis
.Value = "31.12.2022"
End With

With Me.txt_Min_Rate_BME
'.Value = Tabelle5.Range("N2").Value
End With

With Me.txt_Min_Rate_BRO
'.Value = Tabelle5.Range("N3").Value
End With

With Me.txt_bgr_dz
.Value = Tabelle5.Range("K16").Value
End With

With Me.txt_SpielR_BME
'.Value = Tabelle5.Range("K5").Value
End With

With Me.txt_SpielR_BRO
'.Value = Tabelle5.Range("K4").Value
End With

'___________________________Checkbox_________________________

'.........................Page2...........................

With Me.chk_Mo
.Value = True
End With

With Me.chk_Di
.Value = True
End With

With Me.chk_Mi
.Value = True
End With

With Me.chk_Do
.Value = True
End With

With Me.chk_Fr
.Value = True
End With

With Me.chk_Sa
.Value = True
End With

With Me.chk_So
.Value = True
End With

lastrow = Tabelle5.Range("BF1048576").End(xlUp).row
If lastrow = 1 Then lastrow = 2
myarray = Tabelle5.Range("BF2:BU" & lastrow)

'Listbox

With Me.ListBox1
    .Clear
    .ColumnCount = 16
    .ColumnHeads = False
    .ColumnWidths = "80 Pt;60 Pt;70 Pt;70 Pt;75 Pt;75 Pt;70 Pt;70 Pt;75 Pt;70 Pt;75 Pt;70 Pt;70 Pt;75 Pt;75 Pt;75 Pt;75 Pt;75 Pt;75 Pt"
    .List = myarray
    .MultiSelect = 0
End With

End Sub
