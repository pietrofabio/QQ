VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Quick_Quote_v3 
   Caption         =   "Quick Quote"
   ClientHeight    =   7635
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   9228.001
   OleObjectBlob   =   "Quick_Quote_v3.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "Quick_Quote_v3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit                 'v. 3.2.5

Dim bolexit As Boolean
Dim Erster_Aufruf As Boolean

Public Sub cmd_QuickQuote_Click()

Dim dateAnreise As String, dateAbreise As String
Dim dateAnreise2 As String, dateAnreise3 As String, dateAnreise4 As String
Dim i As Integer
Dim intZeileAn As Integer, intzeileAb As Integer
Dim intZeile2 As Integer, intZeile3 As Integer, intZeile4 As Integer
Dim intLGR_Row As Integer, intZeileAnLGR As Integer
Dim intZeile2LGR As Integer, intZeile3LGR As Integer, intZeile4LGR As Integer
Dim rngAnreise As Range, rngAbreise As Range
Dim rngTag2 As Range, rngTag3 As Range, rngTag4 As Range

Dim rngsaison1 As Range, rngsaison2 As Range, rngsaison3 As Range, rngsaison4 As Range
Dim LGR_Zeile As Integer, LGR_Zeile2 As Integer, LGR_Zeile3 As Integer, LGR_Zeile4 As Integer
Dim firstAddress As String
Dim intAnzTage As Integer
Dim Tag As String, Monat As String, Jahr As String, FehlerGrund As String
Dim firstDate As Date, firstDate2 As Date, firstDate3 As Date, firstDate4 As Date

Dim sinBAR1 As Single, sinBAR2 As Single, sinBAR3 As Single, sinBAR4 As Single
Dim sinSELL1 As Single, sinSELL2 As Single, sinSELL3 As Single, sinSELL4 As Single
Dim sinSELLWalk As Single, sinSELLWish As Single
Dim sinSELLWalk1 As Single, sinSELLWish1 As Single
Dim sinSELLWalk2 As Single, sinSELLWish2 As Single
Dim sinSELLWalk3 As Single, sinSELLWish3 As Single
Dim sinSELLWalk4 As Single, sinSELLWish4 As Single

bolexit = False

'******************************************************************************
'******************************************************************************

'************************************
' ******** I. PRÜFUNGEN  ************
'************************************

'       1. Prüfung Datumseingabe
'       2. An- und Abreisedatum formatieren
'       3. Zeilenindex bestimmen
'           i)      * 1 Tag Aufenthalt *
'           ii)     * 2 Tage Aufenthalt *
'           iii)    * 3 Tage Aufenthalt *
'           iv)     * 4 Tage Aufenthalt *
'       4. Abfrage MAX LOS (Aktuell 4 Tage möglich)

'************************************
' ********** II. ABFRAGEN ***********
'************************************

'       1. Abfrage Occupancy
'           i)      -BRO-
'           ii)     -BME-
'           iii)    -LGR-
'               a) '   1 Tag    '
'               b) '   2 Tage    '
'               c) '   3 Tage    '
'               d) '   4 Tage    '

'       2. Abfrage Zimmeranzahl vs. Maxanzahl
'           i)      .Segment BRO.
'           ii)     .Segment BME.
'           iii)    .Segment LGR.

'       3. Abfrage nach MLOS Restriktion im DC
'           i)      .Abfrage Tag 1.
'           ii)     .Abfrage Tag 2.
'           iii)    .Abfrage Tag 3.
'           iv)     .Abfrage Tag 4.

'       4. Abfrage CXL Policy

'       5. Segment LGR Abfrage Budget Room Nights und Wochentage Closeouts

'*************************************
' ****  III. Berechnungsparameter ****
'*************************************

'   IV.  BAR Raten zuweisen
'       1.  - Abfrage Segment BRO -
'       2.  - Abfrage Segment BME -
'           i)  * Plenum Protect Prüfung *
'       3.  - Abfrage Segment LGR -
'           i)      . 1 Tag Aufenthalt .
'           ii)     . 2 Tage Aufenthalt .
'           iii)    . 3 Tage Aufenthalt .
'           iv)     . 4 Tage Aufenthalt .

'************************************
' ********** V. INITIALIZE **********
'************************************

'       1. UserForm_Initialize

'************************************
' ******** VI. LGR PRÜFUNGEN ********
'************************************

'       1. Sub LGR_2020
'           i) . - 1 Tag Aufenthalt - .
'           ii) . - 2 Tage Aufenthalt - .
'           iii) . - 3 Tage Aufenthalt - .
'           iv) . - 4 Tage Aufenthalt - .
'       2. Sub LGR_2021
'           i) . / 1 Tag Aufenthalt / .
'           ii) . / 2 Tage Aufenthalt / .
'           iii) . / 3 Tage Aufenthalt / .
'           iv) . / 4 Tage Aufenthalt / .

'*************************************
' *********** VII. DIVERSE ***********
'*************************************
'       1. cbo_Segment_Change

'*************************************
' *********** VIII. CLICKS ***********
'*************************************

'       1. cmd_EMailErstellen_Click
'       2. cmd_BGR_Konfig_Click
'       3. CMD_LGR_Konfig_Click

'*************************************
' ********** IX. FUNKTIONEN **********
'*************************************

'       1. Function DateiInBearbeitung
'       2. Function OVERLAPS
'       3. Function CellContentCanBeInterpretedAsADate
        
'******************************************************************************
'******************************************************************************
        

'____________________________Prüfung Datumseingabe_________________________________

Me.txt_ZimmerAnzahl.BackColor = rgbWhite
Me.txt_ZimmerAnzahl_dz.BackColor = rgbWhite

With Me.txt_Anreise
    If CStr(Len(Me.txt_Anreise)) = "6" Then
        Tag = Left(Me.txt_Anreise.Value, 2)
        Monat = Mid(Me.txt_Anreise.Value, 3, 2)
        Jahr = Right(Me.txt_Anreise, 2)
        Me.txt_Anreise = DateSerial(CInt(Jahr), CInt(Monat), CInt(Tag))
        .SetFocus
    ElseIf CStr(Len(Me.txt_Anreise)) = "8" Then
        Tag = Left(Me.txt_Anreise.Value, 2)
        Monat = Mid(Me.txt_Anreise.Value, 3, 2)
        Jahr = Right(Me.txt_Anreise, 4)
        Me.txt_Anreise = DateSerial(CInt(Jahr), CInt(Monat), CInt(Tag))
        .SetFocus
    ElseIf Not IsDate(.Text) Then
        MsgBox "Please enter an arrival date. (Dateformat: DDMMYY, DDMMYYYY, DD.MM.YYYY, oder DD/MM/YYYY)"
        .Value = Date + 30
        .SetFocus
        Exit Sub
    End If
End With

With Me.txt_Abreise
    If CStr(Len(Me.txt_Abreise)) = "6" Then
        Tag = Left(Me.txt_Abreise.Value, 2)
        Monat = Mid(Me.txt_Abreise.Value, 3, 2)
        Jahr = Right(Me.txt_Abreise, 2)
        Me.txt_Abreise = DateSerial(CInt(Jahr), CInt(Monat), CInt(Tag))
        .SetFocus
    ElseIf CStr(Len(Me.txt_Abreise)) = "8" Then
        Tag = Left(Me.txt_Abreise.Value, 2)
        Monat = Mid(Me.txt_Abreise.Value, 3, 2)
        Jahr = Right(Me.txt_Abreise, 4)
        Me.txt_Abreise = DateSerial(CInt(Jahr), CInt(Monat), CInt(Tag))
        .SetFocus
    ElseIf Not IsDate(.Text) Then
        MsgBox "Please enter a departure date. (Dateformat: DDMMYY, DDMMYYYY, DD.MM.YYYY, oder DD/MM/YYYY)"
        .Value = Date + 31
        .SetFocus
        Exit Sub
    End If
End With

If CDate(Me.txt_Anreise.Value) < CDate(Now) Or CDate(Me.txt_Anreise.Value) > DateSerial(2024, 12, 31) Then
    MsgBox "The arrival date must be between today and 31.12.2024"
    Me.txt_Anreise.SetFocus
    Exit Sub
End If

If CDate(Me.txt_Abreise.Value) < CDate(Now) + 1 Or CDate(Me.txt_Abreise.Value) > DateSerial(2025, 1, 1) Then
    MsgBox "The departure date must be between tomorrow and 01.01.2025"
    Me.txt_Anreise.SetFocus
    Exit Sub
End If

If CDate(Me.txt_Abreise.Value) <= CDate(Me.txt_Anreise.Value) Then
    MsgBox "Arrival date must be earlier than departure date."
    Exit Sub
End If

If Me.txt_ZimmerAnzahl.Value = "" And Me.txt_ZimmerAnzahl_dz.Value = "" Then
    Me.txt_ZimmerAnzahl.BackColor = rgbPink
    Me.txt_ZimmerAnzahl_dz.BackColor = rgbPink
    Exit Sub
End If

If Me.txt_ZimmerAnzahl.Value = "" And Me.txt_ZimmerAnzahl_dz.Value <> "" Then
    Me.txt_ZimmerAnzahl.Value = 0
End If

If Me.txt_ZimmerAnzahl.Value <> "" And Me.txt_ZimmerAnzahl_dz.Value = "" Then
    Me.txt_ZimmerAnzahl_dz.Value = 0
End If

'____________________________An- und Abreisedatum formatieren____________________


Me.txt_Anreise = Format(Me.txt_Anreise, "DD.MM.YYYY")
Me.txt_Abreise = Format(Me.txt_Abreise, "DD.MM.YYYY")

dateAnreise = Me.txt_Anreise
dateAbreise = Me.txt_Abreise
Tabelle5.Range("I27").Value = ""

'____________________________Zeilenindex bestimmen____________________________

intAnzTage = CDate(dateAbreise) - CDate(dateAnreise)
intLGR_Row = Tabelle5.Range("O1048576").End(xlUp).row ' + 1

Tabelle5.Range("N23").Value = CDate(Me.txt_Anreise.Value)
Tabelle5.Range("N24").Value = CDate(Me.txt_Anreise.Value) + 1
Tabelle5.Range("N25").Value = CDate(Me.txt_Anreise.Value) + 2
Tabelle5.Range("N26").Value = CDate(Me.txt_Anreise.Value) + 3

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

''''''''''''''''''''''''''''''
    Application.Calculate ''''
''''''''''''''''''''''''''''''

Select Case intAnzTage

'........................* 1 Tag Aufenthalt *...................

    Case 1
    With Tabelle18.Range("B:B")
    Set rngAnreise = .Find(what:=CDate(dateAnreise), LookIn:=xlFormulas, lookat:=xlWhole)
    End With
    intZeileAn = rngAnreise.row
    
'........................* 2 Tage Aufenthalt *...................
    
    Case 2
    With Tabelle18.Range("B:B")
    Set rngAnreise = .Find(what:=CDate(dateAnreise), LookIn:=xlFormulas, lookat:=xlWhole)
    End With
    intZeileAn = rngAnreise.row
    
    dateAnreise2 = CDate(dateAnreise) + 1
    With Tabelle18.Range("B:B")
    Set rngTag2 = .Find(what:=CDate(dateAnreise2), LookIn:=xlFormulas, lookat:=xlWhole)
    End With
    intZeile2 = rngTag2.row
    
'........................* 3 Tage Aufenthalt *...................
    
    Case 3
    With Tabelle18.Range("B:B")
    Set rngAnreise = .Find(what:=CDate(dateAnreise), LookIn:=xlFormulas, lookat:=xlWhole)
    End With
    intZeileAn = rngAnreise.row
    
    dateAnreise2 = CDate(dateAnreise) + 1
    With Tabelle18.Range("B:B")
    Set rngTag2 = .Find(what:=CDate(dateAnreise2), LookIn:=xlFormulas, lookat:=xlWhole)
    End With
    intZeile2 = rngTag2.row
    
    dateAnreise3 = CDate(dateAnreise) + 2
    With Tabelle18.Range("B:B")
    Set rngTag3 = .Find(what:=CDate(dateAnreise3), LookIn:=xlFormulas, lookat:=xlWhole)
    End With
    intZeile3 = rngTag3.row
    
'........................* 4 Tage Aufenthalt *...................
    
    Case 4
    With Tabelle18.Range("B:B")
    Set rngAnreise = .Find(what:=CDate(dateAnreise), LookIn:=xlFormulas, lookat:=xlWhole)
    End With
    intZeileAn = rngAnreise.row
    
    dateAnreise2 = CDate(dateAnreise) + 1
    With Tabelle18.Range("B:B")
    Set rngTag2 = .Find(what:=CDate(dateAnreise2), LookIn:=xlFormulas, lookat:=xlWhole)
    End With
    intZeile2 = rngTag2.row
    
    dateAnreise3 = CDate(dateAnreise) + 2
    With Tabelle18.Range("B:B")
    Set rngTag3 = .Find(what:=CDate(dateAnreise3), LookIn:=xlFormulas, lookat:=xlWhole)
    End With
    intZeile3 = rngTag3.row
    
    dateAnreise4 = CDate(dateAnreise) + 3
    With Tabelle18.Range("B:B")
    Set rngTag4 = .Find(what:=CDate(dateAnreise4), LookIn:=xlFormulas, lookat:=xlWhole)
    End With
    intZeile4 = rngTag4.row
    
'--------------------------------Abfrage MAX LOS (Aktuell 4 Tage möglich) ------------------------------------
    
    Case Is > 4
    Me.lbl_Rate.Caption = ""
    Me.lbl_Preisbereich.Caption = ""
    FehlerGrund = "LOS more than 4 nights"
    Tabelle5.Range("I27").Value = FehlerGrund
    
    With Me.lbl_Quote_Info
    .Caption = "The requested LOS is not eligible for Quick Quote. Please contact your Revenue Manager."
    .ForeColor = RGB(255, 0, 0)
    .AutoSize = False
    .WordWrap = True
    End With
    
    Exit Sub

End Select

'-----------------------------------------------------------------------------------------
'---------------------------------Abfrage Occupancy---------------------------------------
'-----------------------------------------------------------------------------------------

'If Me.chk_Total_Occ.Value = True Then

Dim intOCCSpalte As Integer
Dim intOCCZeile As Integer, intOCCZeileAn As Integer, intOCCZeile2 As Integer, intOCCZeile3 As Integer, intOCCZeile4 As Integer
Dim sinOCCmax As Single, sinOCCmax2 As Single
Dim sinOCCAn As Single, sinOCC2 As Single, sinOCC3 As Single, sinOCC4 As Single
Dim dblOCCAn As Double, dblOCC2 As Double, dblOCC3 As Double, dblOCC4 As Double

On Error Resume Next

sinOCCAn = Tabelle3.PivotTables("pt_history_forecast").GetPivotData("Occ_Pct_DEF", "StayDate", rngAnreise.Value) * 100
sinOCC2 = Tabelle3.PivotTables("pt_history_forecast").GetPivotData("Occ_Pct_DEF", "StayDate", rngTag2.Value) * 100
sinOCC3 = Tabelle3.PivotTables("pt_history_forecast").GetPivotData("Occ_Pct_DEF", "StayDate", rngTag3.Value) * 100
sinOCC4 = Tabelle3.PivotTables("pt_history_forecast").GetPivotData("Occ_Pct_DEF", "StayDate", rngTag4.Value) * 100
On Error GoTo 0

'sinOCCmax = Application.Max(sinOCCAn, sinOCC2, sinOCC3, sinOCC4)

On Error Resume Next

firstDate = CDate(Tabelle5.Range("N27"))
firstDate2 = CDate(Tabelle5.Range("N28"))
firstDate3 = CDate(Tabelle5.Range("N29"))
firstDate4 = CDate(Tabelle5.Range("N30"))

''''''''''''''''''''''''''''''' Select Segment ''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Select Case Me.cbo_Segment

'--------------------------------- # BRO # ---------------------------------------

Case Is = "BRO"

    intOCCZeile = Tabelle5.Range("AI2:AI" & Tabelle5.Range("AI1048576").End(xlUp).row).Find(what:=CDate(dateAnreise), LookIn:=xlValues).row
    
    On Error Resume Next
    dblOCCAn = Tabelle5.Cells(intOCCZeile, 48).Value
    dblOCC2 = Tabelle5.Cells(intOCCZeile + 1, 48).Value
    dblOCC3 = Tabelle5.Cells(intOCCZeile + 2, 48).Value
    dblOCC4 = Tabelle5.Cells(intOCCZeile + 3, 48).Value
    On Error GoTo 0
    
    'sinOCCmax2 = Application.Max(dblOCCAn, dblOCC2, dblOCC3, dblOCC4)

'--------------------------------- # BME # ---------------------------------------

Case Is = "BME"

    intOCCZeile = Tabelle5.Range("AI2:AI" & Tabelle5.Range("AI1048576").End(xlUp).row).Find(what:=CDate(dateAnreise), LookIn:=xlValues).row
    
    On Error Resume Next
    dblOCCAn = Tabelle5.Cells(intOCCZeile, 40).Value
    dblOCC2 = Tabelle5.Cells(intOCCZeile + 1, 40).Value
    dblOCC3 = Tabelle5.Cells(intOCCZeile + 2, 40).Value
    dblOCC4 = Tabelle5.Cells(intOCCZeile + 3, 40).Value
    On Error GoTo 0
    
    'sinOCCmax2 = Application.Max(dblOCCAn, dblOCC2, dblOCC3, dblOCC4)

'--------------------------------- # LGR # ---------------------------------------

Case Is = "LGR"

    On Error Resume Next
    intOCCZeileAn = Tabelle5.Range("O1:O" & intLGR_Row).Find(what:=CStr(firstDate), LookIn:=xlValues).row
    intOCCZeile2 = Tabelle5.Range("O1:O" & intLGR_Row).Find(what:=CStr(firstDate2), LookIn:=xlValues).row
    intOCCZeile3 = Tabelle5.Range("O1:O" & intLGR_Row).Find(what:=CStr(firstDate3), LookIn:=xlValues).row
    intOCCZeile4 = Tabelle5.Range("O1:O" & intLGR_Row).Find(what:=CStr(firstDate4), LookIn:=xlValues).row
    If intOCCZeileAn = 0 Then
    intOCCZeileAn = Tabelle5.Range("O1:O" & intLGR_Row).Find(what:=firstDate, LookIn:=xlValues).row
    End If
    If intOCCZeile2 = 0 Then
    intOCCZeile2 = Tabelle5.Range("O1:O" & intLGR_Row).Find(what:=firstDate2, LookIn:=xlValues).row
    End If
    If intOCCZeile3 = 0 Then
    intOCCZeile3 = Tabelle5.Range("O1:O" & intLGR_Row).Find(what:=firstDate3, LookIn:=xlValues).row
    End If
    If intOCCZeile4 = 0 Then
    intOCCZeile4 = Tabelle5.Range("O1:O" & intLGR_Row).Find(what:=firstDate4, LookIn:=xlValues).row
    End If
    On Error GoTo 0
    
    On Error Resume Next
    dblOCCAn = Tabelle5.Cells(intOCCZeileAn, 22).Value
    dblOCC2 = Tabelle5.Cells(intOCCZeile2, 22).Value
    dblOCC3 = Tabelle5.Cells(intOCCZeile3, 22).Value
    dblOCC4 = Tabelle5.Cells(intOCCZeile4, 22).Value
    On Error GoTo 0

Select Case intAnzTage

'   1 Tag

Case Is = 1

    If Tabelle5.Range("N27").Value = "" Then
        Me.lbl_Rate.Caption = ""
        Me.lbl_Preisbereich.Caption = ""
        With Me.lbl_Quote_Info
        .Caption = "There is no season available for the arrival date. Please contact your Revenue Manager."
        .ForeColor = RGB(255, 0, 0)
        .AutoSize = False
        .WordWrap = True
        End With
        bolexit = True
        Exit Sub
    End If

'   2 Tage

Case Is = 2

    If Tabelle5.Range("N27").Value = "" Then
        Me.lbl_Rate.Caption = ""
        Me.lbl_Preisbereich.Caption = ""
        With Me.lbl_Quote_Info
        .Caption = "There is no season available for the arrival date. Please contact your Revenue Manager."
        .ForeColor = RGB(255, 0, 0)
        .AutoSize = False
        .WordWrap = True
        End With
        bolexit = True
        Exit Sub
    ElseIf Tabelle5.Range("N28").Value = "" Then
        Me.lbl_Rate.Caption = ""
        Me.lbl_Preisbereich.Caption = ""
        With Me.lbl_Quote_Info
        .Caption = "There is no season available for the 2nd day. Please contact your Revenue Manager."
        .ForeColor = RGB(255, 0, 0)
        .AutoSize = False
        .WordWrap = True
        End With
        bolexit = True
        Exit Sub
    End If

'   3 Tage

Case Is = 3

    If Tabelle5.Range("N27").Value = "" Then
        Me.lbl_Rate.Caption = ""
        Me.lbl_Preisbereich.Caption = ""
        With Me.lbl_Quote_Info
        .Caption = "There is no season available for the arrival date. Please contact your Revenue Manager."
        .ForeColor = RGB(255, 0, 0)
        .AutoSize = False
        .WordWrap = True
        End With
        bolexit = True
        Exit Sub
    ElseIf Tabelle5.Range("N28").Value = "" Then
        Me.lbl_Rate.Caption = ""
        Me.lbl_Preisbereich.Caption = ""
        With Me.lbl_Quote_Info
        .Caption = "There is no season available for the 2nd day. Please contact your Revenue Manager."
        .ForeColor = RGB(255, 0, 0)
        .AutoSize = False
        .WordWrap = True
        End With
        bolexit = True
        Exit Sub
    ElseIf Tabelle5.Range("N29").Value = "" Then
        Me.lbl_Rate.Caption = ""
        Me.lbl_Preisbereich.Caption = ""
        With Me.lbl_Quote_Info
        .Caption = "There is no season available for the 3rd day. Please contact your Revenue Manager."
        .ForeColor = RGB(255, 0, 0)
        .AutoSize = False
        .WordWrap = True
        End With
        bolexit = True
        Exit Sub
    End If

'   4 Tage

Case Is = 4

    If Tabelle5.Range("N27").Value = "" Then
        Me.lbl_Rate.Caption = ""
        Me.lbl_Preisbereich.Caption = ""
        With Me.lbl_Quote_Info
        .Caption = "There is no season available for the arrival date. Please contact your Revenue Manager."
        .ForeColor = RGB(255, 0, 0)
        .AutoSize = False
        .WordWrap = True
        End With
        bolexit = True
        Exit Sub
    ElseIf Tabelle5.Range("N28").Value = "" Then
        Me.lbl_Rate.Caption = ""
        Me.lbl_Preisbereich.Caption = ""
        With Me.lbl_Quote_Info
        .Caption = "There is no season available for the 2nd day. Please contact your Revenue Manager."
        .ForeColor = RGB(255, 0, 0)
        .AutoSize = False
        .WordWrap = True
        End With
        bolexit = True
        Exit Sub
    ElseIf Tabelle5.Range("N29").Value = "" Then
        Me.lbl_Rate.Caption = ""
        Me.lbl_Preisbereich.Caption = ""
        With Me.lbl_Quote_Info
        .Caption = "There is no season available for the 3rd day. Please contact your Revenue Manager."
        .ForeColor = RGB(255, 0, 0)
        .AutoSize = False
        .WordWrap = True
        End With
        bolexit = True
        Exit Sub
    ElseIf Tabelle5.Range("N30").Value = "" Then
        Me.lbl_Rate.Caption = ""
        Me.lbl_Preisbereich.Caption = ""
        With Me.lbl_Quote_Info
        .Caption = "There is no season available for the 4th day. Please contact your Revenue Manager."
        .ForeColor = RGB(255, 0, 0)
        .AutoSize = False
        .WordWrap = True
        End With
        bolexit = True
        Exit Sub
    End If

End Select

End Select

'''''''''''''''''''''''''''''''' Select Aufenthalt '''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Select Case intAnzTage

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   1 Tag    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Case Is = 1
    
    If sinOCCAn > dblOCCAn Then
        Me.lbl_Rate.Caption = ""
        Me.lbl_Preisbereich.Caption = ""
        FehlerGrund = "Occupancy is too high: " & Round(sinOCCAn, 2) & "% > " & dblOCCAn & " %"
        Tabelle5.Range("I27").Value = FehlerGrund
        With Me.lbl_Quote_Info
        .Caption = "The occupancy on the arrival date exceeds the eligibility of Quick Quote (" & Round(sinOCCAn, 2) & "% > " & Round(dblOCCAn, 2) & "%). Please contact your Revenue Manager."
        .ForeColor = RGB(255, 0, 0)
        .AutoSize = False
        .WordWrap = True
        End With
        Exit Sub
    End If
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   2 Tage    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
Case Is = 2
    
    If sinOCCAn > dblOCCAn Then
        Me.lbl_Rate.Caption = ""
        Me.lbl_Preisbereich.Caption = ""
        FehlerGrund = "Occupancy too high: " & Round(sinOCCAn, 2) & "% > " & dblOCCAn & " %"
        Tabelle5.Range("I27").Value = FehlerGrund
        With Me.lbl_Quote_Info
        .Caption = "The occupancy on the arrival date exceeds the eligibility of Quick Quote  (" & Round(sinOCCAn, 2) & "% > " & Round(dblOCCAn, 2) & "%). Please contact your Revenue Manager."
        .ForeColor = RGB(255, 0, 0)
        .AutoSize = False
        .WordWrap = True
        End With
        Exit Sub
    ElseIf sinOCC2 > dblOCC2 Then
        Me.lbl_Rate.Caption = ""
        Me.lbl_Preisbereich.Caption = ""
        FehlerGrund = "Occupancy too high: " & Round(sinOCC2, 2) & "% > " & dblOCC2 & " %"
        Tabelle5.Range("I27").Value = FehlerGrund
        With Me.lbl_Quote_Info
        .Caption = "The occupancy on the 2nd day exceeds the eligibility of Quick Quote  (" & Round(sinOCC2, 2) & "% > " & Round(dblOCC2, 2) & "%). Please contact your Revenue Manager."
        .ForeColor = RGB(255, 0, 0)
        .AutoSize = False
        .WordWrap = True
        End With
        Exit Sub
    End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   3 Tage    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Case Is = 3
    
    If sinOCCAn > dblOCCAn Then
        Me.lbl_Rate.Caption = ""
        Me.lbl_Preisbereich.Caption = ""
        FehlerGrund = "Occupancy too high: " & Round(sinOCCAn, 2) & "% > " & dblOCCAn & " %"
        Tabelle5.Range("I27").Value = FehlerGrund
        With Me.lbl_Quote_Info
        .Caption = "The occupancy on the arrival date exceeds the eligibility of Quick Quote  (" & Round(sinOCCAn, 2) & "% > " & Round(dblOCCAn, 2) & "%). Please contact your Revenue Manager."
        .ForeColor = RGB(255, 0, 0)
        .AutoSize = False
        .WordWrap = True
        End With
        Exit Sub
    ElseIf sinOCC2 > dblOCC2 Then
        Me.lbl_Rate.Caption = ""
        Me.lbl_Preisbereich.Caption = ""
        FehlerGrund = "Occupancy too high: " & Round(sinOCC2, 2) & "% > " & dblOCC2 & " %"
        Tabelle5.Range("I27").Value = FehlerGrund
        With Me.lbl_Quote_Info
        .Caption = "The occupancy on the 2nd day exceeds the eligibility of Quick Quote  (" & Round(sinOCC2, 2) & "% > " & Round(dblOCC2, 2) & "%). Please contact your Revenue Manager."
        .ForeColor = RGB(255, 0, 0)
        .AutoSize = False
        .WordWrap = True
        End With
        Exit Sub
    ElseIf sinOCC3 > dblOCC3 Then
        Me.lbl_Rate.Caption = ""
        Me.lbl_Preisbereich.Caption = ""
        FehlerGrund = "Occupancy too high: " & Round(sinOCC3, 2) & "% > " & dblOCC3 & " %"
        Tabelle5.Range("I27").Value = FehlerGrund
        With Me.lbl_Quote_Info
        .Caption = "The occupancy on the 3rd day exceeds the eligibility of Quick Quote  (" & Round(sinOCC3, 2) & "% > " & Round(dblOCC3, 2) & "%). Please contact your Revenue Manager."
        .ForeColor = RGB(255, 0, 0)
        .AutoSize = False
        .WordWrap = True
        End With
        Exit Sub
    End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   4 Tage    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Case Is = 4
    
    If sinOCCAn > dblOCCAn Then
        Me.lbl_Rate.Caption = ""
        Me.lbl_Preisbereich.Caption = ""
        FehlerGrund = "Occupancy too high: " & Round(sinOCCAn, 2) & "% > " & dblOCCAn & " %"
        Tabelle5.Range("I27").Value = FehlerGrund
        With Me.lbl_Quote_Info
        .Caption = "The occupancy on the arrival date exceeds the eligibility of Quick Quote  (" & Round(sinOCCAn, 2) & "% > " & Round(dblOCCAn, 2) & "%). Please contact your Revenue Manager."
        .ForeColor = RGB(255, 0, 0)
        .AutoSize = False
        .WordWrap = True
        End With
        Exit Sub
    ElseIf sinOCC2 > dblOCC2 Then
        Me.lbl_Rate.Caption = ""
        Me.lbl_Preisbereich.Caption = ""
        FehlerGrund = "Occupancy too high: " & Round(sinOCC2, 2) & "% > " & dblOCC2 & " %"
        Tabelle5.Range("I27").Value = FehlerGrund
        With Me.lbl_Quote_Info
        .Caption = "The occupancy on the 2nd day exceeds the eligibility of Quick Quote  (" & Round(sinOCC2, 2) & "% > " & Round(dblOCC2, 2) & "%). Please contact your Revenue Manager."
        .ForeColor = RGB(255, 0, 0)
        .AutoSize = False
        .WordWrap = True
        End With
        Exit Sub
    ElseIf sinOCC3 > dblOCC3 Then
        Me.lbl_Rate.Caption = ""
        Me.lbl_Preisbereich.Caption = ""
        FehlerGrund = "Occupancy too high: " & Round(sinOCC3, 2) & "% > " & dblOCC3 & " %"
        Tabelle5.Range("I27").Value = FehlerGrund
        With Me.lbl_Quote_Info
        .Caption = "The occupancy on the 3rd day exceeds the eligibility of Quick Quote  (" & Round(sinOCC3, 2) & "% > " & Round(dblOCC3, 2) & "%). Please contact your Revenue Manager."
        .ForeColor = RGB(255, 0, 0)
        .AutoSize = False
        .WordWrap = True
        End With
        Exit Sub
    ElseIf sinOCC4 > dblOCC4 Then
        Me.lbl_Rate.Caption = ""
        Me.lbl_Preisbereich.Caption = ""
        FehlerGrund = "Occupancy too high: " & Round(sinOCC4, 2) & "% > " & dblOCC4 & " %"
        Tabelle5.Range("I27").Value = FehlerGrund
        With Me.lbl_Quote_Info
        .Caption = "The occupancy on the 4th day exceeds the eligibility of Quick Quote  (" & Round(sinOCC4, 2) & "% > " & Round(dblOCC4, 2) & "%). Please contact your Revenue Manager."
        .ForeColor = RGB(255, 0, 0)
        .AutoSize = False
        .WordWrap = True
        End With
        Exit Sub
    End If

End Select

'End If

'------------------------------------------------------------------------------------------------------
'------------------------------- Abfrage Zimmeranzahl vs. Maxanzahl -----------------------------------
'------------------------------------------------------------------------------------------------------

Dim intMaxRoomBro         As Integer
Dim intMaxRoomBme         As Integer
Dim intMaxRoomLgr1        As Integer
Dim intMaxRoomLgr2        As Integer
Dim intMaxRoomLgr3        As Integer
Dim intMaxRoomLgr4        As Integer
Dim intEZ                 As Integer
Dim intDZ                 As Integer
Dim rowDiscount           As Integer

'intMaxRoomBro = Tabelle5.Range("L8").Value
'intMaxRoomBme = Tabelle5.Range("L15").Value
'intMaxRoomLgr = Tabelle5.Range("W2").Value


'..................................Segment BRO..................................

If Me.cbo_Segment.Value = "BRO" Then

    intEZ = Me.txt_ZimmerAnzahl.Value * 1
    intDZ = Me.txt_ZimmerAnzahl_dz.Value * 1
    rowDiscount = Tabelle5.Range("AI2:AI" & Tabelle5.Range("AI1048576").End(xlUp).row).Find(what:=rngAnreise.Value, LookIn:=xlFormulas, lookat:=xlWhole).row
    
    Select Case intAnzTage
    
'........................1 Tag Aufenthalt...................
    
    Case 1
            If intEZ + intDZ > Tabelle5.Range("AT" & rowDiscount) Then
                Me.lbl_Rate.Caption = ""
                Me.lbl_Preisbereich.Caption = ""
                FehlerGrund = "The requested no. of rooms on the arrival date exceed BRO Max. No. of Rooms (# " & Tabelle5.Range("AT" & rowDiscount) & ")."
                Tabelle5.Range("I27").Value = FehlerGrund
                
                With Me.lbl_Quote_Info
                    .Caption = FehlerGrund & " Please contact your Revenue Manager."
                    .ForeColor = RGB(255, 0, 0)
                    .AutoSize = False
                    .WordWrap = True
                End With
            Exit Sub
            End If
    
'........................2 Tage Aufenthalt...................
    
    Case 2
        For i = 0 To 1
            If intEZ + intDZ > Tabelle5.Range("AT" & rowDiscount + i) Then
                Me.lbl_Rate.Caption = ""
                Me.lbl_Preisbereich.Caption = ""
                Select Case i
                Case 0
                FehlerGrund = "The requested no. of rooms on the arrival date exceed BRO Max. No. of Rooms (# " & Tabelle5.Range("AT" & rowDiscount) & ")."
                Case 1
                FehlerGrund = "The requested no. of rooms on the 2nd day exceed BRO Max. No. of Rooms (# " & Tabelle5.Range("AT" & rowDiscount) & ")."
                End Select
                Tabelle5.Range("I27").Value = FehlerGrund
                
                With Me.lbl_Quote_Info
                    .Caption = FehlerGrund & " Please contact your Revenue Manager."
                    .ForeColor = RGB(255, 0, 0)
                    .AutoSize = False
                    .WordWrap = True
                End With
            Exit Sub
            End If
        Next i
    
'........................3 Tage Aufenthalt...................
    
    Case 3
        For i = 0 To 2
            If intEZ + intDZ > Tabelle5.Range("AT" & rowDiscount + i) Then
                Me.lbl_Rate.Caption = ""
                Me.lbl_Preisbereich.Caption = ""
                Select Case i
                Case 0
                FehlerGrund = "The requested no. of rooms on the arrival date exceed BRO Max. No. of Rooms (# " & Tabelle5.Range("AT" & rowDiscount) & ")."
                Case 1
                FehlerGrund = "The requested no. of rooms on the 2nd day exceed BRO Max. No. of Rooms (# " & Tabelle5.Range("AT" & rowDiscount) & ")."
                Case 2
                FehlerGrund = "The requested no. of rooms on the 3rd day exceed BRO Max. No. of Rooms (# " & Tabelle5.Range("AT" & rowDiscount) & ")."
                End Select
                Tabelle5.Range("I27").Value = FehlerGrund
                
                With Me.lbl_Quote_Info
                    .Caption = FehlerGrund & " Please contact your Revenue Manager."
                    .ForeColor = RGB(255, 0, 0)
                    .AutoSize = False
                    .WordWrap = True
                End With
            Exit Sub
            End If
        Next i
    
'........................4 Tage Aufenthalt...................
    
    Case 4
        For i = 0 To 3
            If intEZ + intDZ > Tabelle5.Range("AT" & rowDiscount + i) Then
                Me.lbl_Rate.Caption = ""
                Me.lbl_Preisbereich.Caption = ""
                Select Case i
                Case 0
                FehlerGrund = "The requested no. of rooms on the arrival date exceed BRO Max. No. of Rooms (# " & Tabelle5.Range("AT" & rowDiscount) & ")."
                Case 1
                FehlerGrund = "The requested no. of rooms on the 2nd day exceed BRO Max. No. of Rooms (# " & Tabelle5.Range("AT" & rowDiscount) & ")."
                Case 2
                FehlerGrund = "The requested no. of rooms on the 3rd day exceed BRO Max. No. of Rooms (# " & Tabelle5.Range("AT" & rowDiscount) & ")."
                Case 3
                FehlerGrund = "The requested no. of rooms on the 4th day exceed BRO Max. No. of Rooms (# " & Tabelle5.Range("AT" & rowDiscount) & ")."
                End Select
                Tabelle5.Range("I27").Value = FehlerGrund
                With Me.lbl_Quote_Info
                    .Caption = FehlerGrund & " Please contact your Revenue Manager."
                    .ForeColor = RGB(255, 0, 0)
                    .AutoSize = False
                    .WordWrap = True
                End With
            Exit Sub
            End If
        Next i
    End Select
End If

'..................................Segment BME..................................

If Me.cbo_Segment.Value = "BME" Then

    intEZ = Me.txt_ZimmerAnzahl.Value * 1
    intDZ = Me.txt_ZimmerAnzahl_dz.Value * 1
    rowDiscount = Tabelle5.Range("AI2:AI" & Tabelle5.Range("AI1048576").End(xlUp).row).Find(what:=rngAnreise.Value, LookIn:=xlFormulas, lookat:=xlWhole).row
    
    Select Case intAnzTage
    
'........................1 Tag Aufenthalt...................
    
    Case 1
            If intEZ + intDZ > Tabelle5.Range("AL" & rowDiscount) Then
                Me.lbl_Rate.Caption = ""
                Me.lbl_Preisbereich.Caption = ""
                FehlerGrund = "The requested no. of rooms on the arrival date exceed BME Max. No. of Rooms (# " & Tabelle5.Range("AL" & rowDiscount) & ")."
                Tabelle5.Range("I27").Value = FehlerGrund
                
                With Me.lbl_Quote_Info
                    .Caption = FehlerGrund & " Please contact your Revenue Manager."
                    .ForeColor = RGB(255, 0, 0)
                    .AutoSize = False
                    .WordWrap = True
                End With
            Exit Sub
            End If
    
'........................2 Tage Aufenthalt...................
    
    Case 2
        For i = 0 To 1
            If intEZ + intDZ > Tabelle5.Range("AL" & rowDiscount + i) Then
                Me.lbl_Rate.Caption = ""
                Me.lbl_Preisbereich.Caption = ""
                Select Case i
                Case 0
                FehlerGrund = "The requested no. of rooms on the arrival date exceed BME Max. No. of Rooms (# " & Tabelle5.Range("AL" & rowDiscount) & ")."
                Case 1
                FehlerGrund = "The requested no. of rooms on the 2nd day exceed BME Max. No. of Rooms (# " & Tabelle5.Range("AL" & rowDiscount) & ")."
                End Select
                Tabelle5.Range("I27").Value = FehlerGrund
                
                With Me.lbl_Quote_Info
                    .Caption = FehlerGrund & " Please contact your Revenue Manager."
                    .ForeColor = RGB(255, 0, 0)
                    .AutoSize = False
                    .WordWrap = True
                End With
            Exit Sub
            End If
        Next i
    
'........................3 Tage Aufenthalt...................
    
    Case 3
        For i = 0 To 2
            If intEZ + intDZ > Tabelle5.Range("AL" & rowDiscount + i) Then
                Me.lbl_Rate.Caption = ""
                Me.lbl_Preisbereich.Caption = ""
                Select Case i
                Case 0
                FehlerGrund = "The requested no. of rooms on the arrival date exceed BME Max. No. of Rooms (# " & Tabelle5.Range("AL" & rowDiscount) & ")."
                Case 1
                FehlerGrund = "The requested no. of rooms on the 2nd day exceed BME Max. No. of Rooms (# " & Tabelle5.Range("AL" & rowDiscount) & ")."
                Case 2
                FehlerGrund = "The requested no. of rooms on the 3rd day exceed BME Max. No. of Rooms (# " & Tabelle5.Range("AL" & rowDiscount) & ")."
                End Select
                Tabelle5.Range("I27").Value = FehlerGrund
                
                With Me.lbl_Quote_Info
                    .Caption = FehlerGrund & " Please contact your Revenue Manager."
                    .ForeColor = RGB(255, 0, 0)
                    .AutoSize = False
                    .WordWrap = True
                End With
            Exit Sub
            End If
        Next i
    
'........................4 Tage Aufenthalt...................
    
    Case 4
        For i = 0 To 3
            If intEZ + intDZ > Tabelle5.Range("AL" & rowDiscount + i) Then
                Me.lbl_Rate.Caption = ""
                Me.lbl_Preisbereich.Caption = ""
                Select Case i
                Case 0
                FehlerGrund = "The requested no. of rooms on the arrival date exceed BME Max. No. of Rooms (# " & Tabelle5.Range("AL" & rowDiscount) & ")."
                Case 1
                FehlerGrund = "The requested no. of rooms on the 2nd day exceed BME Max. No. of Rooms (# " & Tabelle5.Range("AL" & rowDiscount) & ")."
                Case 2
                FehlerGrund = "The requested no. of rooms on the 3rd day exceed BME Max. No. of Rooms (# " & Tabelle5.Range("AL" & rowDiscount) & ")."
                Case 3
                FehlerGrund = "The requested no. of rooms on the 4th day exceed BME Max. No. of Rooms (# " & Tabelle5.Range("AL" & rowDiscount) & ")."
                End Select
                Tabelle5.Range("I27").Value = FehlerGrund
                
                With Me.lbl_Quote_Info
                    .Caption = FehlerGrund & " Please contact your Revenue Manager."
                    .ForeColor = RGB(255, 0, 0)
                    .AutoSize = False
                    .WordWrap = True
                End With
            Exit Sub
            End If
        Next i
    End Select
End If


'..................................Segment LGR..................................

If Me.cbo_Segment.Value = "LGR" Then

    On Error Resume Next
    intOCCZeileAn = Tabelle5.Range("O1:O" & intLGR_Row).Find(what:=CStr(firstDate), LookIn:=xlValues).row
    intOCCZeile2 = Tabelle5.Range("O1:O" & intLGR_Row).Find(what:=CStr(firstDate2), LookIn:=xlValues).row
    intOCCZeile3 = Tabelle5.Range("O1:O" & intLGR_Row).Find(what:=CStr(firstDate3), LookIn:=xlValues).row
    intOCCZeile4 = Tabelle5.Range("O1:O" & intLGR_Row).Find(what:=CStr(firstDate4), LookIn:=xlValues).row
    If intOCCZeileAn = 0 Then
    intOCCZeileAn = Tabelle5.Range("O1:O" & intLGR_Row).Find(what:=firstDate, LookIn:=xlValues).row
    End If
    If intOCCZeile2 = 0 Then
    intOCCZeile2 = Tabelle5.Range("O1:O" & intLGR_Row).Find(what:=firstDate2, LookIn:=xlValues).row
    End If
    If intOCCZeile3 = 0 Then
    intOCCZeile3 = Tabelle5.Range("O1:O" & intLGR_Row).Find(what:=firstDate3, LookIn:=xlValues).row
    End If
    If intOCCZeile4 = 0 Then
    intOCCZeile4 = Tabelle5.Range("O1:O" & intLGR_Row).Find(what:=firstDate4, LookIn:=xlValues).row
    End If
    On Error GoTo 0

    intEZ = Me.txt_ZimmerAnzahl.Value * 1
    intDZ = Me.txt_ZimmerAnzahl_dz.Value * 1

    Select Case intAnzTage
    
        Case 1
        
        intMaxRoomLgr1 = Tabelle5.Range("W" & intOCCZeileAn).Value
        
        If intEZ + intDZ > intMaxRoomLgr1 Then
            Me.lbl_Rate.Caption = ""
            Me.lbl_Preisbereich.Caption = ""
            FehlerGrund = "The requested no. of rooms on the arrival date exceed LGR Max. No. of Rooms (# " & intMaxRoomLgr1 & ")."
            Tabelle5.Range("I27").Value = FehlerGrund
            
            With Me.lbl_Quote_Info
                .Caption = FehlerGrund & " Please contact your Revenue Manager."
                .ForeColor = RGB(255, 0, 0)
                .AutoSize = False
                .WordWrap = True
            End With
        Exit Sub
        End If
        
        Case 2
        
        intMaxRoomLgr1 = Tabelle5.Range("W" & intOCCZeileAn).Value
        intMaxRoomLgr2 = Tabelle5.Range("W" & intOCCZeile2).Value
        
        If intEZ + intDZ > intMaxRoomLgr1 Then
            Me.lbl_Rate.Caption = ""
            Me.lbl_Preisbereich.Caption = ""
            FehlerGrund = "The requested no. of rooms on the arrival date exceed LGR Max. No. of Rooms (# " & intMaxRoomLgr1 & ")."
            Tabelle5.Range("I27").Value = FehlerGrund
            
            With Me.lbl_Quote_Info
                .Caption = FehlerGrund & " Please contact your Revenue Manager."
                .ForeColor = RGB(255, 0, 0)
                .AutoSize = False
                .WordWrap = True
            End With
        Exit Sub
        ElseIf intEZ + intDZ > intMaxRoomLgr2 Then
            Me.lbl_Rate.Caption = ""
            Me.lbl_Preisbereich.Caption = ""
            FehlerGrund = "The requested no. of rooms on the 2nd day exceed LGR Max. No. of Rooms (# " & intMaxRoomLgr2 & ")."
            Tabelle5.Range("I27").Value = FehlerGrund
            
            With Me.lbl_Quote_Info
                .Caption = FehlerGrund & " Please contact your Revenue Manager."
                .ForeColor = RGB(255, 0, 0)
                .AutoSize = False
                .WordWrap = True
            End With
        Exit Sub
        End If
        
        Case 3
        
        intMaxRoomLgr1 = Tabelle5.Range("W" & intOCCZeileAn).Value
        intMaxRoomLgr2 = Tabelle5.Range("W" & intOCCZeile2).Value
        intMaxRoomLgr3 = Tabelle5.Range("W" & intOCCZeile3).Value
        
        If intEZ + intDZ > intMaxRoomLgr1 Then
            Me.lbl_Rate.Caption = ""
            Me.lbl_Preisbereich.Caption = ""
            FehlerGrund = "The requested no. of rooms on the arrival date exceed LGR Max. No. of Rooms (# " & intMaxRoomLgr1 & ")."
            Tabelle5.Range("I27").Value = FehlerGrund
            
            With Me.lbl_Quote_Info
                .Caption = FehlerGrund & " Please contact your Revenue Manager."
                .ForeColor = RGB(255, 0, 0)
                .AutoSize = False
                .WordWrap = True
            End With
        Exit Sub
        ElseIf intEZ + intDZ > intMaxRoomLgr2 Then
            Me.lbl_Rate.Caption = ""
            Me.lbl_Preisbereich.Caption = ""
            FehlerGrund = "The requested no. of rooms on the 2nd day exceed LGR Max. No. of Rooms (# " & intMaxRoomLgr2 & ")."
            Tabelle5.Range("I27").Value = FehlerGrund
            
            With Me.lbl_Quote_Info
                .Caption = FehlerGrund & " Please contact your Revenue Manager."
                .ForeColor = RGB(255, 0, 0)
                .AutoSize = False
                .WordWrap = True
            End With
        Exit Sub
        ElseIf intEZ + intDZ > intMaxRoomLgr3 Then
            Me.lbl_Rate.Caption = ""
            Me.lbl_Preisbereich.Caption = ""
            FehlerGrund = "The requested no. of rooms on the 3rd day exceed LGR Max. No. of Rooms (# " & intMaxRoomLgr3 & ")."
            Tabelle5.Range("I27").Value = FehlerGrund
            
            With Me.lbl_Quote_Info
                .Caption = FehlerGrund & " Please contact your Revenue Manager."
                .ForeColor = RGB(255, 0, 0)
                .AutoSize = False
                .WordWrap = True
            End With
        Exit Sub
        End If
        
        Case 4
        
        intMaxRoomLgr1 = Tabelle5.Range("W" & intOCCZeileAn).Value
        intMaxRoomLgr2 = Tabelle5.Range("W" & intOCCZeile2).Value
        intMaxRoomLgr3 = Tabelle5.Range("W" & intOCCZeile3).Value
        intMaxRoomLgr4 = Tabelle5.Range("W" & intOCCZeile4).Value
        
        If intEZ + intDZ > intMaxRoomLgr1 Then
            Me.lbl_Rate.Caption = ""
            Me.lbl_Preisbereich.Caption = ""
            FehlerGrund = "The requested no. of rooms on the arrival date exceed LGR Max. No. of Rooms (# " & intMaxRoomLgr1 & ")."
            Tabelle5.Range("I27").Value = FehlerGrund
            
            With Me.lbl_Quote_Info
                .Caption = FehlerGrund & " Please contact your Revenue Manager."
                .ForeColor = RGB(255, 0, 0)
                .AutoSize = False
                .WordWrap = True
            End With
        Exit Sub
        ElseIf intEZ + intDZ > intMaxRoomLgr2 Then
            Me.lbl_Rate.Caption = ""
            Me.lbl_Preisbereich.Caption = ""
            FehlerGrund = "The requested no. of rooms on the 2nd day exceed LGR Max. No. of Rooms (# " & intMaxRoomLgr2 & ")."
            Tabelle5.Range("I27").Value = FehlerGrund
            
            With Me.lbl_Quote_Info
                .Caption = FehlerGrund & " Please contact your Revenue Manager."
                .ForeColor = RGB(255, 0, 0)
                .AutoSize = False
                .WordWrap = True
            End With
        Exit Sub
        ElseIf intEZ + intDZ > intMaxRoomLgr3 Then
            Me.lbl_Rate.Caption = ""
            Me.lbl_Preisbereich.Caption = ""
            FehlerGrund = "The requested no. of rooms on the 3rd day exceed LGR Max. No. of Rooms (# " & intMaxRoomLgr3 & ")."
            Tabelle5.Range("I27").Value = FehlerGrund
            
            With Me.lbl_Quote_Info
                .Caption = FehlerGrund & " Please contact your Revenue Manager."
                .ForeColor = RGB(255, 0, 0)
                .AutoSize = False
                .WordWrap = True
            End With
        Exit Sub
        ElseIf intEZ + intDZ > intMaxRoomLgr4 Then
            Me.lbl_Rate.Caption = ""
            Me.lbl_Preisbereich.Caption = ""
            FehlerGrund = "The requested no. of rooms on the 4th day exceed LGR Max. No. of Rooms (# " & intMaxRoomLgr4 & ")."
            Tabelle5.Range("I27").Value = FehlerGrund
            
            With Me.lbl_Quote_Info
                .Caption = FehlerGrund & " Please contact your Revenue Manager."
                .ForeColor = RGB(255, 0, 0)
                .AutoSize = False
                .WordWrap = True
            End With
        Exit Sub
        End If
    
    End Select

End If


'-------------------------------------------------------------------------------
'......................Abfrage nach MLOS Restriktion im DC......................
'-------------------------------------------------------------------------------


'.................Abfrage Tag 1......................

Dim intMLOSAn As String, intMLOS2 As String, intMLOS3 As String, intMLOS4 As String

Select Case intAnzTage

Case 1

    intMLOSAn = Tabelle18.Cells(intZeileAn, 26).Value
    If intMLOSAn = "" Then
        intMLOSAn = 0
    End If
        If chk_Dyn_MLOS = False And intMLOSAn > 0 Then      ' Dynamische MLOS aus, MLOS eingestellt --> QQ nicht möglich
            Select Case intMLOSAn
                Case Is <> ""
                Me.lbl_Rate.Caption = ""
                Me.lbl_Preisbereich.Caption = ""
                FehlerGrund = "Min LOS " & intMLOSAn & " is set up"
                Tabelle5.Range("I27").Value = FehlerGrund
                    With Me.lbl_Quote_Info
                        .Caption = "Due to an MLOS " & intMLOSAn & " restriction on the arrival date, not eligible for Quick Quote. Please contact your Revenue Manager."
                        .ForeColor = RGB(255, 0, 0)
                        .AutoSize = False
                        .WordWrap = True
                    End With
                Exit Sub
            End Select
        ElseIf chk_Dyn_MLOS = True Then
            Select Case intMLOSAn
                Case Is > intAnzTage
                Me.lbl_Rate.Caption = ""
                Me.lbl_Preisbereich.Caption = ""
                FehlerGrund = "Min LOS " & intMLOSAn & " is set up"
                Tabelle5.Range("I27").Value = FehlerGrund

                    With Me.lbl_Quote_Info
                        .Caption = "Due to an MLOS " & intMLOSAn & " restriction on the arrival date, not eligible for Quick Quote. Please contact your Revenue Manager."
                        .ForeColor = RGB(255, 0, 0)
                        .AutoSize = False
                        .WordWrap = True
                    End With
                Exit Sub
            End Select
        End If
'.................Abfrage Tag 2......................

Case 2

For i = 0 To 1
    Select Case i
        Case 0
        intMLOS2 = Tabelle18.Cells(intZeileAn, 26).Value
        Case 1
        intMLOS2 = Tabelle18.Cells(intZeile2, 26).Value
    End Select
    If intMLOS2 = "" Then
        intMLOS2 = 0
    End If
        If chk_Dyn_MLOS = False And intMLOS2 > 0 Then      ' Dynamische MLOS aus, MLOS eingestellt --> QQ nicht möglich
            Select Case intMLOS2
                Case Is <> ""
                Me.lbl_Rate.Caption = ""
                Me.lbl_Preisbereich.Caption = ""
                FehlerGrund = "Min LOS " & intMLOS2 & " is set up"
                Tabelle5.Range("I27").Value = FehlerGrund
                    With Me.lbl_Quote_Info
                        .Caption = "Due to an MLOS " & intMLOS2 & " restriction on the 2nd day, not eligible for Quick Quote. Please contact your Revenue Manager."
                        .ForeColor = RGB(255, 0, 0)
                        .AutoSize = False
                        .WordWrap = True
                    End With
                Exit Sub
            End Select
        ElseIf chk_Dyn_MLOS = True Then
            Select Case intMLOS2
                Case Is > intAnzTage
                Me.lbl_Rate.Caption = ""
                Me.lbl_Preisbereich.Caption = ""
                FehlerGrund = "Min LOS " & intMLOS2 & " is set up"
                Tabelle5.Range("I27").Value = FehlerGrund
                    With Me.lbl_Quote_Info
                        .Caption = "Due to an MLOS " & intMLOS2 & " restriction on the 2nd day, not eligible for Quick Quote. Please contact your Revenue Manager."
                        .ForeColor = RGB(255, 0, 0)
                        .AutoSize = False
                        .WordWrap = True
                    End With
                Exit Sub
            End Select
        End If
Next i

'.................Abfrage Tag 3......................

Case 3

For i = 0 To 2
    Select Case i
        Case 0
        intMLOS3 = Tabelle18.Cells(intZeileAn, 26).Value
        Case 1
        intMLOS3 = Tabelle18.Cells(intZeile2, 26).Value
        Case 2
        intMLOS3 = Tabelle18.Cells(intZeile3, 26).Value
    End Select
    If intMLOS3 = "" Then
        intMLOS3 = 0
    End If
        If chk_Dyn_MLOS = False And intMLOS3 > 0 Then      ' Dynamische MLOS aus, MLOS eingestellt --> QQ nicht möglich
            Select Case intMLOS3
                Case Is <> ""
                Me.lbl_Rate.Caption = ""
                Me.lbl_Preisbereich.Caption = ""
                FehlerGrund = "Min LOS " & intMLOS3 & " is set up"
                Tabelle5.Range("I27").Value = FehlerGrund
                    With Me.lbl_Quote_Info
                        .Caption = "Due to an MLOS " & intMLOS3 & " restriction on the 3rd day, not eligible for Quick Quote. Please contact your Revenue Manager."
                        .ForeColor = RGB(255, 0, 0)
                        .AutoSize = False
                        .WordWrap = True
                    End With
                Exit Sub
            End Select
        ElseIf chk_Dyn_MLOS = True Then
            Select Case intMLOS3
                Case Is > intAnzTage
                Me.lbl_Rate.Caption = ""
                Me.lbl_Preisbereich.Caption = ""
                FehlerGrund = "Min LOS " & intMLOS3 & " is set up"
                Tabelle5.Range("I27").Value = FehlerGrund
                    With Me.lbl_Quote_Info
                        .Caption = "Due to an MLOS " & intMLOS3 & " restriction on the 3rd day, not eligible for Quick Quote. Please contact your Revenue Manager."
                        .ForeColor = RGB(255, 0, 0)
                        .AutoSize = False
                        .WordWrap = True
                    End With
                Exit Sub
            End Select
        End If
Next i

'.................Abfrage Tag 4......................

Case 4

For i = 0 To 3
    Select Case i
        Case 0
        intMLOS4 = Tabelle18.Cells(intZeileAn, 26).Value
        Case 1
        intMLOS4 = Tabelle18.Cells(intZeile2, 26).Value
        Case 2
        intMLOS4 = Tabelle18.Cells(intZeile3, 26).Value
        Case 3
        intMLOS4 = Tabelle18.Cells(intZeile4, 26).Value
    End Select
    If intMLOS4 = "" Then
        intMLOS4 = 0
    End If
        If chk_Dyn_MLOS = False And intMLOS4 > 0 Then      ' Dynamische MLOS aus, MLOS eingestellt --> QQ nicht möglich
            Select Case intMLOS4
                Case Is <> ""
                Me.lbl_Rate.Caption = ""
                Me.lbl_Preisbereich.Caption = ""
                FehlerGrund = "Min LOS " & intMLOS4 & " is set up"
                Tabelle5.Range("I27").Value = FehlerGrund
                    With Me.lbl_Quote_Info
                        .Caption = "Due to an MLOS " & intMLOS4 & " restriction on the 4th day, not eligible for Quick Quote. Please contact your Revenue Manager."
                        .ForeColor = RGB(255, 0, 0)
                        .AutoSize = False
                        .WordWrap = True
                    End With
                Exit Sub
            End Select
        ElseIf chk_Dyn_MLOS = True Then
            Select Case intMLOS4
                Case Is > intAnzTage
                Me.lbl_Rate.Caption = ""
                Me.lbl_Preisbereich.Caption = ""
                FehlerGrund = "Min LOS " & intMLOS4 & " is set up"
                Tabelle5.Range("I27").Value = FehlerGrund
                    With Me.lbl_Quote_Info
                        .Caption = "Due to an MLOS " & intMLOS2 & " restriction on the 4th day, not eligible for Quick Quote. Please contact your Revenue Manager."
                        .ForeColor = RGB(255, 0, 0)
                        .AutoSize = False
                        .WordWrap = True
                    End With
                Exit Sub
            End Select
        End If
Next i

End Select

'--------------------------------------------------------------------------------------
'--------------------------------Abfrage CXL Policy------------------------------------
'--------------------------------------------------------------------------------------

If Me.chk_CXL_Policy = True Then

    Dim strCXLAn As String, strCXL2 As String, strCXL3 As String, strCXL4 As String
    
    '.................Abfrage Tag 1......................
    
    strCXLAn = Tabelle18.Cells(intZeileAn, 24).Value
    
    Select Case strCXLAn
    
    Case Is <> ""
        Me.lbl_Rate.Caption = ""
        Me.lbl_Preisbereich.Caption = ""
        FehlerGrund = "CXL " & strCXLAn & " is set up"
        Tabelle5.Range("I27").Value = FehlerGrund
        With Me.lbl_Quote_Info
        .Caption = "Due to a CXL policy on the arrival date, not eligible for Quick Quote. Please contact your Revenue Manager."
        .ForeColor = RGB(255, 0, 0)
        .AutoSize = False
        .WordWrap = True
        End With
        Exit Sub
    End Select
    
    '.................Abfrage Tag 2......................
    
    On Error Resume Next
    strCXL2 = Tabelle18.Cells(intZeile2, 24).Value
    
    Select Case strCXL2
    
    Case Is <> ""
        Me.lbl_Rate.Caption = ""
        Me.lbl_Preisbereich.Caption = ""
        FehlerGrund = "CXL " & strCXL2 & " is set up"
        Tabelle5.Range("I27").Value = FehlerGrund
        With Me.lbl_Quote_Info
        .Caption = "Due to a CXL policy on the 2nd day, not eligible for Quick Quote. Please contact your Revenue Manager."
        .ForeColor = RGB(255, 0, 0)
        .AutoSize = False
        .WordWrap = True
        End With
        Exit Sub
    End Select
    
    '.................Abfrage Tag 3......................
    
    strCXL3 = Tabelle18.Cells(intZeile3, 24).Value
    
    Select Case strCXL3
    
    Case Is <> ""
        Me.lbl_Rate.Caption = ""
        Me.lbl_Preisbereich.Caption = ""
        FehlerGrund = "CXL " & strCXL3 & " is set up"
        Tabelle5.Range("I27").Value = FehlerGrund
        With Me.lbl_Quote_Info
        .Caption = "Due to a CXL policy on the 3rd day, not eligible for Quick Quote. Please contact your Revenue Manager."
        .ForeColor = RGB(255, 0, 0)
        .AutoSize = False
        .WordWrap = True
        End With
        Exit Sub
    End Select
    
    '.................Abfrage Tag 4......................
    
    strCXL4 = Tabelle18.Cells(intZeile4, 24).Value
    
    Select Case strCXL4
    
    Case Is <> ""
        Me.lbl_Rate.Caption = ""
        Me.lbl_Preisbereich.Caption = ""
        FehlerGrund = "CXL " & strCXL4 & " is set up"
        Tabelle5.Range("I27").Value = FehlerGrund
        With Me.lbl_Quote_Info
        .Caption = "Due to a CXL policy on the 4th day, not eligible for Quick Quote. Please contact your Revenue Manager."
        .ForeColor = RGB(255, 0, 0)
        .AutoSize = False
        .WordWrap = True
        End With
        Exit Sub
    End Select

End If

On Error GoTo 0

'-----------------------------------------------------------------------------------------------------------------------
'____________________________Segment LGR Abfrage Wochentage Closeouts____________________________
'-----------------------------------------------------------------------------------------------------------------------

Select Case Me.cbo_Segment

Case "LGR"

    Application.Calculation = xlCalculationManual

    Dim intSaison1   As Integer
    Dim intSaison2   As Integer
    Dim intSaison3   As Integer
    Dim intSaison4   As Integer

    Call LGR_2021
    
    If bolexit = True Then
        Exit Sub
    End If


End Select

'__________________________Berechnungsparameter______________________________________

Dim sinBRODiscount_ As Object
Dim sinBMEDiscount_ As Object
Dim sinBME_Fix_ As Object
Dim sinBRO_Fix_ As Object
Dim sinBME_Min_ As Object
Dim sinBRO_Min_ As Object
Dim sinSpielR_BRO_ As Object
Dim sinSpielR_BME_ As Object
Dim sinSpielR_LGR_ As Object
Dim sin_Kat_ As Object

Dim endPreis As Single
Dim sinBPSplit As Single
Dim sinBgrDZ As Single
Dim sinSpielR_BME As Single
Dim sinSpielR_BRO As Single
Dim sin_Kat_avg As Single

'sinSpielR_BRO = Tabelle5.Range("L4").Value
'sinSpielR_BME = Tabelle5.Range("L5").Value
'sinSpielR_LGR = Tabelle5.Range("L6").Value
sinBPSplit = Tabelle5.Range("K7").Value
sinBgrDZ = Tabelle5.Range("K16").Value

If Me.txt_ZimmerAnzahl.Value = "" Then
    Me.txt_ZimmerAnzahl.Value = 0
End If

If Me.txt_ZimmerAnzahl_dz.Value = "" Then
    Me.txt_ZimmerAnzahl_dz.Value = 0
End If


'**************************************************************************
'********************** BAR Raten zuweisen ********************************
'**************************************************************************


On Error Resume Next
sinBAR1 = Tabelle18.Range("K" & intZeileAn).Value
sinBAR2 = Tabelle18.Range("K" & intZeile2).Value
sinBAR3 = Tabelle18.Range("K" & intZeile3).Value
sinBAR4 = Tabelle18.Range("K" & intZeile4).Value

'--------------------------------Ratenberechung-------------------------------------------
Select Case Me.cbo_Segment

    Dim sinSellWishK As Single, sinSellWalkK As Single
    Dim sinSellWishK1 As Single, sinSellWalkK1 As Single
    Dim sinSellWishK2 As Single, sinSellWalkK2 As Single
    Dim sinSellWishK3 As Single, sinSellWalkK3 As Single
    Dim sinSellWishK4 As Single, sinSellWalkK4 As Single

'-------------------------------- Abfrage Segment BRO ---------------------------------------

Case "BRO"

    Set sinBRODiscount_ = CreateObject("Scripting.Dictionary")
    Set sinBRO_Fix_ = CreateObject("Scripting.Dictionary")
    Set sinBRO_Min_ = CreateObject("Scripting.Dictionary")
    Set sinSpielR_BRO_ = CreateObject("Scripting.Dictionary")
    Set sin_Kat_ = CreateObject("Scripting.Dictionary")
    
    For i = 0 To intAnzTage - 1
        sinBRODiscount_(i + 1) = Tabelle5.Range("AS" & rowDiscount + i).Value
        sinBRO_Fix_(i + 1) = Tabelle5.Range("AU" & rowDiscount + i).Value
        sinBRO_Min_(i + 1) = Tabelle5.Range("AW" & rowDiscount + i).Value
        sinSpielR_BRO_(i + 1) = Tabelle5.Range("AX" & rowDiscount + i).Value
        sin_Kat_(i + 1) = Tabelle5.Range("AR" & rowDiscount + i).Value
    Next i
    
'------------CHECK FIX AND MIN-------------
    
    If sinBRO_Fix_(1) <> 0 Then
    sinSELL1 = sinBRO_Fix_(1) * (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value)
    ElseIf sinBRO_Min_(1) > sinBAR1 - sinBAR1 * sinBRODiscount_(1) Then
    sinSELL1 = sinBRO_Min_(1) * (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value)
    Else
    sinSELL1 = (sinBAR1 - sinBAR1 * sinBRODiscount_(1)) * (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value)
    End If
    
    If sinBRO_Fix_(2) <> 0 Then
    sinSELL2 = sinBRO_Fix_(2) * (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value)
    ElseIf sinBRO_Min_(2) > sinBAR2 - sinBAR2 * sinBRODiscount_(2) Then
    sinSELL2 = sinBRO_Min_(2) * (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value)
    Else
    sinSELL2 = (sinBAR2 - sinBAR2 * sinBRODiscount_(2)) * (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value)
    End If
    
    If sinBRO_Fix_(3) <> 0 Then
    sinSELL3 = sinBRO_Fix_(3) * (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value)
    ElseIf sinBRO_Min_(3) > sinBAR3 - sinBAR3 * sinBRODiscount_(3) Then
    sinSELL3 = sinBRO_Min_(3) * (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value)
    Else
    sinSELL3 = (sinBAR3 - sinBAR3 * sinBRODiscount_(3)) * (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value)
    End If
    
    If sinBRO_Fix_(4) <> 0 Then
    sinSELL4 = sinBRO_Fix_(4) * (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value)
    ElseIf sinBRO_Min_(4) > sinBAR4 - sinBAR4 * sinBRODiscount_(4) Then
    sinSELL4 = sinBRO_Min_(4) * (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value)
    Else
    sinSELL4 = (sinBAR4 - sinBAR4 * sinBRODiscount_(4)) * (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value)
    End If

'-------------------LOS CHECK--------------------

    Select Case intAnzTage
    
    Case 1
    
        If sinBRO_Fix_(1) <> 0 Then
        
            sinSELLWish1 = sinBRO_Fix_(1)
            
        ElseIf sinBRO_Min_(1) > sinBAR1 - sinBAR1 * sinBRODiscount_(1) Then
        
            sinSELLWish1 = sinBRO_Min_(1)
            
        Else
            
            If sinSELL1 = 0 Then
                
                Me.lbl_Rate.Caption = ""
                Me.lbl_Preisbereich.Caption = ""
                FehlerGrund = "BAR Rate not available"
                Tabelle5.Range("I27").Value = FehlerGrund
                With Me.lbl_Quote_Info
                .Caption = "Due to the lack of a BAR Rate, not eligible for Quick Quote. Please contact your Revenue Manager."
                .ForeColor = RGB(255, 0, 0)
                .AutoSize = False
                .WordWrap = True
                End With
                Exit Sub
                
            Else
                
                sinSELLWish1 = sinSELL1 / (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value) + sinBPSplit
                sinSELLWalk1 = sinSELL1 / (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value) - sinSpielR_BRO_(1) + sinBPSplit
                
            End If
        
        End If
        
        sinSellWishK = sinSELLWish1 '+ (sinSELLWish / 100 * Me.txt_Kommission)
        sinSellWalkK = sinSELLWalk1 '+ (sinSELLWalk / 100 * Me.txt_Kommission)
        
    Case 2
    
    'Anreise
    
        If sinBRO_Fix_(1) <> 0 Then
        
            sinSELLWish1 = sinBRO_Fix_(1)
            
        ElseIf sinBRO_Min_(1) > sinBAR1 - sinBAR1 * sinBRODiscount_(1) Then
        
            sinSELLWish1 = sinBRO_Min_(1)
            
        Else
        
            If sinSELL1 = 0 Then
                
                Me.lbl_Rate.Caption = ""
                Me.lbl_Preisbereich.Caption = ""
                FehlerGrund = "BAR Rate not available"
                Tabelle5.Range("I27").Value = FehlerGrund
                With Me.lbl_Quote_Info
                .Caption = "Due to the lack of a BAR Rate on the arrival date, not eligible for Quick Quote. Please contact your Revenue Manager."
                .ForeColor = RGB(255, 0, 0)
                .AutoSize = False
                .WordWrap = True
                End With
                Exit Sub
                
            Else
                
                sinSELLWish1 = sinSELL1 / (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value) + sinBPSplit
                sinSELLWalk1 = sinSELL1 / (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value) - sinSpielR_BRO_(1) + sinBPSplit
                
            End If
        
        End If
        
        sinSellWishK1 = sinSELLWish1 '+ (sinSELLWish / 100 * Me.txt_Kommission)
        sinSellWalkK1 = sinSELLWalk1 '+ (sinSELLWalk / 100 * Me.txt_Kommission)
        
    'Tag 2
        
        If sinBRO_Fix_(2) <> 0 Then
        
            sinSELLWish2 = sinBRO_Fix_(2)
            
        ElseIf sinBRO_Min_(2) > sinBAR2 - sinBAR2 * sinBRODiscount_(2) Then
        
            sinSELLWish2 = sinBRO_Min_(2)
            
        Else
            
            If sinSELL2 = 0 Then
                
                Me.lbl_Rate.Caption = ""
                Me.lbl_Preisbereich.Caption = ""
                FehlerGrund = "BAR Rate not available"
                Tabelle5.Range("I27").Value = FehlerGrund
                With Me.lbl_Quote_Info
                .Caption = "Due to the lack of a BAR Rate on Day 2, not eligible for Quick Quote. Please contact your Revenue Manager."
                .ForeColor = RGB(255, 0, 0)
                .AutoSize = False
                .WordWrap = True
                End With
                Exit Sub
                
            Else
                
                sinSELLWish2 = sinSELL2 / (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value) + sinBPSplit
                sinSELLWalk2 = sinSELL2 / (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value) - sinSpielR_BRO_(2) + sinBPSplit
                
            End If
                        
        End If
        
        sinSellWishK2 = sinSELLWish2 '+ (sinSELLWish / 100 * Me.txt_Kommission)
        sinSellWalkK2 = sinSELLWalk2 '+ (sinSELLWalk / 100 * Me.txt_Kommission)
            
        sinSellWishK = (sinSellWishK1 + sinSellWishK2) / intAnzTage
        sinSellWalkK = (sinSellWalkK1 + sinSellWalkK2) / intAnzTage
        
    Case 3
    
    'Anreise
    
        If sinBRO_Fix_(1) <> 0 Then
            
            sinSELLWish1 = sinBRO_Fix_(1)
            
        ElseIf sinBRO_Min_(1) > sinBAR1 - sinBAR1 * sinBRODiscount_(1) Then
        
            sinSELLWish1 = sinBRO_Min_(1)
            
        Else
            
            If sinSELL1 = 0 Then
                
                Me.lbl_Rate.Caption = ""
                Me.lbl_Preisbereich.Caption = ""
                FehlerGrund = "BAR Rate not available"
                Tabelle5.Range("I27").Value = FehlerGrund
                With Me.lbl_Quote_Info
                .Caption = "Due to the lack of a BAR Rate on the arrival date, not eligible for Quick Quote. Please contact your Revenue Manager."
                .ForeColor = RGB(255, 0, 0)
                .AutoSize = False
                .WordWrap = True
                End With
                Exit Sub
                
            Else
                
                sinSELLWish1 = sinSELL1 / (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value) + sinBPSplit
                sinSELLWalk1 = sinSELL1 / (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value) - sinSpielR_BRO_(1) + sinBPSplit
            
            End If
        
        End If
        
        sinSellWishK1 = sinSELLWish1 '+ (sinSELLWish / 100 * Me.txt_Kommission)
        sinSellWalkK1 = sinSELLWalk1 '+ (sinSELLWalk / 100 * Me.txt_Kommission)
        
    'Tag 2
        
        If sinBRO_Fix_(2) <> 0 Then
            
            sinSELLWish2 = sinBRO_Fix_(2)
            
        ElseIf sinBRO_Min_(2) > sinBAR2 - sinBAR2 * sinBRODiscount_(2) Then
        
            sinSELLWish2 = sinBRO_Min_(2)
            
        Else
            
            If sinSELL2 = 0 Then
                
                Me.lbl_Rate.Caption = ""
                Me.lbl_Preisbereich.Caption = ""
                FehlerGrund = "BAR Rate not available"
                Tabelle5.Range("I27").Value = FehlerGrund
                With Me.lbl_Quote_Info
                .Caption = "Due to the lack of a BAR Rate on Day 2, not eligible for Quick Quote. Please contact your Revenue Manager."
                .ForeColor = RGB(255, 0, 0)
                .AutoSize = False
                .WordWrap = True
                End With
                Exit Sub
                
            Else
                
                sinSELLWish2 = sinSELL2 / (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value) + sinBPSplit
                sinSELLWalk2 = sinSELL2 / (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value) - sinSpielR_BRO_(2) + sinBPSplit
            
            End If
                    
        End If
        
        sinSellWishK2 = sinSELLWish2 '+ (sinSELLWish / 100 * Me.txt_Kommission)
        sinSellWalkK2 = sinSELLWalk2 '+ (sinSELLWalk / 100 * Me.txt_Kommission)
        
    'Tag 3
        
        If sinBRO_Fix_(3) <> 0 Then
            
            sinSELLWish3 = sinBRO_Fix_(3)
            
        ElseIf sinBRO_Min_(3) > sinBAR3 - sinBAR3 * sinBRODiscount_(3) Then
        
            sinSELLWish3 = sinBRO_Min_(3)
            
        Else
            
            If sinSELL3 = 0 Then
                
                Me.lbl_Rate.Caption = ""
                Me.lbl_Preisbereich.Caption = ""
                FehlerGrund = "BAR Rate not available"
                Tabelle5.Range("I27").Value = FehlerGrund
                With Me.lbl_Quote_Info
                .Caption = "Due to the lack of a BAR Rate on Day 3, not eligible for Quick Quote. Please contact your Revenue Manager."
                .ForeColor = RGB(255, 0, 0)
                .AutoSize = False
                .WordWrap = True
                End With
                Exit Sub
                
            Else
                
                sinSELLWish3 = sinSELL3 / (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value) + sinBPSplit
                sinSELLWalk3 = sinSELL3 / (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value) - sinSpielR_BRO_(3) + sinBPSplit
            
            End If
        
        End If
        
        sinSellWishK3 = sinSELLWish3 '+ (sinSELLWish / 100 * Me.txt_Kommission)
        sinSellWalkK3 = sinSELLWalk3 '+ (sinSELLWalk / 100 * Me.txt_Kommission)
            
        sinSellWishK = (sinSellWishK1 + sinSellWishK2 + sinSellWishK3) / intAnzTage
        sinSellWalkK = (sinSellWalkK1 + sinSellWalkK2 + sinSellWalkK3) / intAnzTage
        
    Case 4
    
    'Anreise
    
        If sinBRO_Fix_(1) <> 0 Then
            
            sinSELLWish1 = sinBRO_Fix_(1)
            
        ElseIf sinBRO_Min_(1) > sinBAR1 - sinBAR1 * sinBRODiscount_(1) Then
        
            sinSELLWish1 = sinBRO_Min_(1)
            
        Else
            
            If sinSELL1 = 0 Then
                
                Me.lbl_Rate.Caption = ""
                Me.lbl_Preisbereich.Caption = ""
                FehlerGrund = "BAR Rate not available"
                Tabelle5.Range("I27").Value = FehlerGrund
                With Me.lbl_Quote_Info
                .Caption = "Due to the lack of a BAR Rate on the arrival date, not eligible for Quick Quote. Please contact your Revenue Manager."
                .ForeColor = RGB(255, 0, 0)
                .AutoSize = False
                .WordWrap = True
                End With
                Exit Sub
                
            Else
                
                sinSELLWish1 = sinSELL1 / (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value) + sinBPSplit
                sinSELLWalk1 = sinSELL1 / (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value) - sinSpielR_BRO_(1) + sinBPSplit
            
            End If
        
        End If
        
        sinSellWishK1 = sinSELLWish1 '+ (sinSELLWish / 100 * Me.txt_Kommission)
        sinSellWalkK1 = sinSELLWalk1 '+ (sinSELLWalk / 100 * Me.txt_Kommission)
        
    'Tag 2
        
        If sinBRO_Fix_(2) <> 0 Then
            
            sinSELLWish2 = sinBRO_Fix_(2)
            
        ElseIf sinBRO_Min_(2) > sinBAR2 - sinBAR2 * sinBRODiscount_(2) Then
        
            sinSELLWish2 = sinBRO_Min_(2)
            
        Else
            
            If sinSELL2 = 0 Then
                
                Me.lbl_Rate.Caption = ""
                Me.lbl_Preisbereich.Caption = ""
                FehlerGrund = "BAR Rate not available"
                Tabelle5.Range("I27").Value = FehlerGrund
                With Me.lbl_Quote_Info
                .Caption = "Due to the lack of a BAR Rate on Day 2, not eligible for Quick Quote. Please contact your Revenue Manager."
                .ForeColor = RGB(255, 0, 0)
                .AutoSize = False
                .WordWrap = True
                End With
                Exit Sub
                
            Else
                
                sinSELLWish2 = sinSELL2 / (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value) + sinBPSplit
                sinSELLWalk2 = sinSELL2 / (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value) - sinSpielR_BRO_(2) + sinBPSplit
            
            End If
                    
        End If
        
        sinSellWishK2 = sinSELLWish2 '+ (sinSELLWish / 100 * Me.txt_Kommission)
        sinSellWalkK2 = sinSELLWalk2 '+ (sinSELLWalk / 100 * Me.txt_Kommission)
        
    'Tag 3
        
        If sinBRO_Fix_(3) <> 0 Then
            
            sinSELLWish3 = sinBRO_Fix_(3)
            
        ElseIf sinBRO_Min_(3) > sinBAR3 - sinBAR3 * sinBRODiscount_(3) Then
        
            sinSELLWish3 = sinBRO_Min_(3)
            
        Else
            
            If sinSELL3 = 0 Then
                
                Me.lbl_Rate.Caption = ""
                Me.lbl_Preisbereich.Caption = ""
                FehlerGrund = "BAR Rate not available"
                Tabelle5.Range("I27").Value = FehlerGrund
                With Me.lbl_Quote_Info
                .Caption = "Due to the lack of a BAR Rate on Day 3, not eligible for Quick Quote. Please contact your Revenue Manager."
                .ForeColor = RGB(255, 0, 0)
                .AutoSize = False
                .WordWrap = True
                End With
                Exit Sub
                
            Else
                
                sinSELLWish3 = sinSELL3 / (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value) + sinBPSplit
                sinSELLWalk3 = sinSELL3 / (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value) - sinSpielR_BRO_(3) + sinBPSplit
            
            End If
        
        End If
        
        sinSellWishK3 = sinSELLWish3 '+ (sinSELLWish / 100 * Me.txt_Kommission)
        sinSellWalkK3 = sinSELLWalk3 '+ (sinSELLWalk / 100 * Me.txt_Kommission)
            
    'Tag 4
        
        If sinBRO_Fix_(4) <> 0 Then
            
            sinSELLWish4 = sinBRO_Fix_(4)
            
        ElseIf sinBRO_Min_(4) > sinBAR4 - sinBAR4 * sinBRODiscount_(4) Then
        
            sinSELLWish4 = sinBRO_Min_(4)
            
        Else
            
            If sinSELL4 = 0 Then
                
                Me.lbl_Rate.Caption = ""
                Me.lbl_Preisbereich.Caption = ""
                FehlerGrund = "BAR Rate not available"
                Tabelle5.Range("I27").Value = FehlerGrund
                With Me.lbl_Quote_Info
                .Caption = "Due to the lack of a BAR Rate on Day 4, not eligible for Quick Quote. Please contact your Revenue Manager."
                .ForeColor = RGB(255, 0, 0)
                .AutoSize = False
                .WordWrap = True
                End With
                Exit Sub
                
            Else
                
                sinSELLWish4 = sinSELL4 / (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value) + sinBPSplit
                sinSELLWalk4 = sinSELL4 / (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value) - sinSpielR_BRO_(4) + sinBPSplit
            
            End If
        
        End If
        
        sinSellWishK4 = sinSELLWish4 '+ (sinSELLWish / 100 * Me.txt_Kommission)
        sinSellWalkK4 = sinSELLWalk4 '+ (sinSELLWalk / 100 * Me.txt_Kommission)
            
        sinSellWishK = (sinSellWishK1 + sinSellWishK2 + sinSellWishK3 + sinSellWishK4) / intAnzTage
        sinSellWalkK = (sinSellWalkK1 + sinSellWalkK2 + sinSellWalkK3 + sinSellWalkK4) / intAnzTage
        
    End Select


'***********BERECHNUNG*************


    endPreis = Application.WorksheetFunction.Max(sinSellWishK, sinSellWalkK)
    
    sin_Kat_avg = (sin_Kat_(1) + sin_Kat_(2) + sin_Kat_(3) + sin_Kat_(4)) / intAnzTage
    
    sinSpielR_BRO = (sinSpielR_BRO_(1) + sinSpielR_BRO_(2) + sinSpielR_BRO_(3) + sinSpielR_BRO_(4)) / intAnzTage
    
    With Me.lbl_Preisbereich
    .Caption = "The price for the group:"
    End With
    
    With Me.lbl_Rate
    .Caption = Round(endPreis, 1) & " €"
    End With
    
    With Me.lbl_Quote_Info
    .Caption = "The prices include breakfast in single room." & Chr(10) & _
                "The category surcharge for the residence is " & sin_Kat_avg & " €." & Chr(10) & _
                "The surcharge for a double room is " & BGR_Config_v3.txt_bgr_dz & " €." & Chr(10) & _
                "The room for negotiation is " & sinSpielR_BRO & " €."
    .ForeColor = RGB(0, 0, 0)
    .AutoSize = False
    .WordWrap = True
    End With

'----------------------------------- Abfrage Segment BME ------------------------------------

Case "BME"

    Set sinBMEDiscount_ = CreateObject("Scripting.Dictionary")
    Set sinBME_Fix_ = CreateObject("Scripting.Dictionary")
    Set sinBME_Min_ = CreateObject("Scripting.Dictionary")
    Set sinSpielR_BME_ = CreateObject("Scripting.Dictionary")
    Set sin_Kat_ = CreateObject("Scripting.Dictionary")
    
    For i = 0 To intAnzTage - 1
        sinBMEDiscount_(i + 1) = Tabelle5.Range("AK" & rowDiscount + i).Value
        sinBME_Fix_(i + 1) = Tabelle5.Range("AM" & rowDiscount + i).Value
        sinBME_Min_(i + 1) = Tabelle5.Range("AO" & rowDiscount + i).Value
        sinSpielR_BME_(i + 1) = Tabelle5.Range("AP" & rowDiscount + i).Value
        sin_Kat_(i + 1) = Tabelle5.Range("AR" & rowDiscount + i).Value
    Next i
    
'------------CHECK FIX AND MIN-------------
    
    If sinBME_Fix_(1) <> 0 Then
    sinSELL1 = sinBME_Fix_(1) * (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value)
    ElseIf sinBME_Min_(1) > sinBAR1 - sinBAR1 * sinBMEDiscount_(1) Then
    sinSELL1 = sinBME_Min_(1) * (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value)
    Else
    sinSELL1 = (sinBAR1 - sinBAR1 * sinBMEDiscount_(1)) * (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value)
    End If
    
    If sinBME_Fix_(2) <> 0 Then
    sinSELL2 = sinBME_Fix_(2) * (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value)
    ElseIf sinBME_Min_(2) > sinBAR2 - sinBAR2 * sinBMEDiscount_(2) Then
    sinSELL2 = sinBME_Min_(2) * (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value)
    Else
    sinSELL2 = (sinBAR2 - sinBAR2 * sinBMEDiscount_(2)) * (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value)
    End If
    
    If sinBME_Fix_(3) <> 0 Then
    sinSELL3 = sinBME_Fix_(3) * (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value)
    ElseIf sinBME_Min_(3) > sinBAR3 - sinBAR3 * sinBMEDiscount_(3) Then
    sinSELL3 = sinBME_Min_(3) * (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value)
    Else
    sinSELL3 = (sinBAR3 - sinBAR3 * sinBMEDiscount_(3)) * (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value)
    End If
    
    If sinBME_Fix_(4) <> 0 Then
    sinSELL4 = sinBME_Fix_(4) * (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value)
    ElseIf sinBME_Min_(4) > sinBAR4 - sinBAR4 * sinBMEDiscount_(4) Then
    sinSELL4 = sinBME_Min_(4) * (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value)
    Else
    sinSELL4 = (sinBAR4 - sinBAR4 * sinBMEDiscount_(4)) * (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value)
    End If
    
'-------------------LOS CHECK--------------------
    
    Select Case intAnzTage
    
    Case 1
    
        If sinBME_Fix_(1) > sinBME_Min_(1) Then
        
            sinSELLWish1 = sinBME_Fix_(1)
            
        ElseIf sinBME_Min_(1) > sinBAR1 - sinBAR1 * sinBMEDiscount_(1) Then
        
            sinSELLWish1 = sinBME_Min_(1)
            
        Else
            
            If sinSELL1 = 0 Then
                
                Me.lbl_Rate.Caption = ""
                Me.lbl_Preisbereich.Caption = ""
                FehlerGrund = "BAR Rate not available"
                Tabelle5.Range("I27").Value = FehlerGrund
                With Me.lbl_Quote_Info
                .Caption = "Due to the lack of a BAR Rate, not eligible for Quick Quote. Please contact your Revenue Manager."
                .ForeColor = RGB(255, 0, 0)
                .AutoSize = False
                .WordWrap = True
                End With
                Exit Sub
                
            Else
                
                sinSELLWish1 = sinSELL1 / (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value) + sinBPSplit
                sinSELLWalk1 = sinSELL1 / (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value) - sinSpielR_BME_(1) + sinBPSplit
                
            End If
        
        End If
        
        sinSellWishK = sinSELLWish1 '+ (sinSELLWish / 100 * Me.txt_Kommission)
        sinSellWalkK = sinSELLWalk1 '+ (sinSELLWalk / 100 * Me.txt_Kommission)

        
    Case 2
    
    'Anreise
    
        If sinBME_Fix_(1) > sinBME_Min_(1) Then
        
            sinSELLWish1 = sinBME_Fix_(1)
            
        ElseIf sinBME_Min_(1) > sinBAR1 - sinBAR1 * sinBMEDiscount_(1) Then
        
            sinSELLWish1 = sinBME_Min_(1)
            
        Else
        
            If sinSELL1 = 0 Then
                
                Me.lbl_Rate.Caption = ""
                Me.lbl_Preisbereich.Caption = ""
                FehlerGrund = "BAR Rate not available"
                Tabelle5.Range("I27").Value = FehlerGrund
                With Me.lbl_Quote_Info
                .Caption = "Due to the lack of a BAR Rate on the arrival date, not eligible for Quick Quote. Please contact your Revenue Manager."
                .ForeColor = RGB(255, 0, 0)
                .AutoSize = False
                .WordWrap = True
                End With
                Exit Sub
                
            Else
                
                sinSELLWish1 = sinSELL1 / (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value) + sinBPSplit
                sinSELLWalk1 = sinSELL1 / (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value) - sinSpielR_BME_(1) + sinBPSplit
                
            End If
        
        End If
        
        sinSellWishK1 = sinSELLWish1 '+ (sinSELLWish / 100 * Me.txt_Kommission)
        sinSellWalkK1 = sinSELLWalk1 '+ (sinSELLWalk / 100 * Me.txt_Kommission)
        
    'Tag 2
        
        If sinBME_Fix_(2) > sinBME_Min_(2) Then
        
            sinSELLWish2 = sinBME_Fix_(2)
            
        ElseIf sinBME_Min_(2) > sinBAR2 - sinBAR2 * sinBMEDiscount_(2) Then
        
            sinSELLWish2 = sinBME_Min_(2)
            
        Else
            
            If sinSELL2 = 0 Then
                
                Me.lbl_Rate.Caption = ""
                Me.lbl_Preisbereich.Caption = ""
                FehlerGrund = "BAR Rate not available"
                Tabelle5.Range("I27").Value = FehlerGrund
                With Me.lbl_Quote_Info
                .Caption = "Due to the lack of a BAR Rate on Day 2, not eligible for Quick Quote. Please contact your Revenue Manager."
                .ForeColor = RGB(255, 0, 0)
                .AutoSize = False
                .WordWrap = True
                End With
                Exit Sub
                
            Else
                
                sinSELLWish2 = sinSELL2 / (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value) + sinBPSplit
                sinSELLWalk2 = sinSELL2 / (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value) - sinSpielR_BME_(2) + sinBPSplit
                
            End If
                        
        End If
        
        sinSellWishK2 = sinSELLWish2 '+ (sinSELLWish / 100 * Me.txt_Kommission)
        sinSellWalkK2 = sinSELLWalk2 '+ (sinSELLWalk / 100 * Me.txt_Kommission)
            
        sinSellWishK = (sinSellWishK1 + sinSellWishK2) / intAnzTage
        sinSellWalkK = (sinSellWalkK1 + sinSellWalkK2) / intAnzTage
        
    Case 3
    
    'Anreise
    
        If sinBME_Fix_(1) > sinBME_Min_(1) Then
            
            sinSELLWish1 = sinBME_Fix_(1)
            
        ElseIf sinBME_Min_(1) > sinBAR1 - sinBAR1 * sinBMEDiscount_(1) Then
        
            sinSELLWish1 = sinBME_Min_(1)
            
        Else
            
            If sinSELL1 = 0 Then
                
                Me.lbl_Rate.Caption = ""
                Me.lbl_Preisbereich.Caption = ""
                FehlerGrund = "BAR Rate not available"
                Tabelle5.Range("I27").Value = FehlerGrund
                With Me.lbl_Quote_Info
                .Caption = "Due to the lack of a BAR Rate on the arrival date, not eligible for Quick Quote. Please contact your Revenue Manager."
                .ForeColor = RGB(255, 0, 0)
                .AutoSize = False
                .WordWrap = True
                End With
                Exit Sub
                
            Else
                
                sinSELLWish1 = sinSELL1 / (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value) + sinBPSplit
                sinSELLWalk1 = sinSELL1 / (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value) - sinSpielR_BME_(1) + sinBPSplit
            
            End If
        
        End If
        
        sinSellWishK1 = sinSELLWish1 '+ (sinSELLWish / 100 * Me.txt_Kommission)
        sinSellWalkK1 = sinSELLWalk1 '+ (sinSELLWalk / 100 * Me.txt_Kommission)
        
    'Tag 2
        
        If sinBME_Fix_(2) > sinBME_Min_(2) Then
            
            sinSELLWish2 = sinBME_Fix_(2)
            
        ElseIf sinBME_Min_(2) > sinBAR2 - sinBAR2 * sinBMEDiscount_(2) Then
        
            sinSELLWish2 = sinBME_Min_(2)
            
        Else
            
            If sinSELL2 = 0 Then
                
                Me.lbl_Rate.Caption = ""
                Me.lbl_Preisbereich.Caption = ""
                FehlerGrund = "BAR Rate not available"
                Tabelle5.Range("I27").Value = FehlerGrund
                With Me.lbl_Quote_Info
                .Caption = "Due to the lack of a BAR Rate on Day 2, not eligible for Quick Quote. Please contact your Revenue Manager."
                .ForeColor = RGB(255, 0, 0)
                .AutoSize = False
                .WordWrap = True
                End With
                Exit Sub
                
            Else
                
                sinSELLWish2 = sinSELL2 / (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value) + sinBPSplit
                sinSELLWalk2 = sinSELL2 / (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value) - sinSpielR_BME_(2) + sinBPSplit
            
            End If
                    
        End If
        
        sinSellWishK2 = sinSELLWish2 '+ (sinSELLWish / 100 * Me.txt_Kommission)
        sinSellWalkK2 = sinSELLWalk2 '+ (sinSELLWalk / 100 * Me.txt_Kommission)
        
    'Tag 3
        
        If sinBME_Fix_(3) > sinBME_Min_(3) Then
            
            sinSELLWish3 = sinBME_Fix_(3)
            
        ElseIf sinBME_Min_(3) > sinBAR3 - sinBAR3 * sinBMEDiscount_(3) Then
        
            sinSELLWish3 = sinBME_Min_(3)
            
        Else
            
            If sinSELL3 = 0 Then
                
                Me.lbl_Rate.Caption = ""
                Me.lbl_Preisbereich.Caption = ""
                FehlerGrund = "BAR Rate not available"
                Tabelle5.Range("I27").Value = FehlerGrund
                With Me.lbl_Quote_Info
                .Caption = "Due to the lack of a BAR Rate on Day 3, not eligible for Quick Quote. Please contact your Revenue Manager."
                .ForeColor = RGB(255, 0, 0)
                .AutoSize = False
                .WordWrap = True
                End With
                Exit Sub
                
            Else
                
                sinSELLWish3 = sinSELL3 / (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value) + sinBPSplit
                sinSELLWalk3 = sinSELL3 / (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value) - sinSpielR_BME_(3) + sinBPSplit
            
            End If
        
        End If
        
        sinSellWishK3 = sinSELLWish3 '+ (sinSELLWish / 100 * Me.txt_Kommission)
        sinSellWalkK3 = sinSELLWalk3 '+ (sinSELLWalk / 100 * Me.txt_Kommission)
            
        sinSellWishK = (sinSellWishK1 + sinSellWishK2 + sinSellWishK3) / intAnzTage
        sinSellWalkK = (sinSellWalkK1 + sinSellWalkK2 + sinSellWalkK3) / intAnzTage
        
    Case 4
    
    'Anreise
    
        If sinBME_Fix_(1) > sinBME_Min_(1) Then
            
            sinSELLWish1 = sinBME_Fix_(1)
            
        ElseIf sinBME_Min_(1) > sinBAR1 - sinBAR1 * sinBMEDiscount_(1) Then
        
            sinSELLWish1 = sinBME_Min_(1)
            
        Else
            
            If sinSELL1 = 0 Then
                
                Me.lbl_Rate.Caption = ""
                Me.lbl_Preisbereich.Caption = ""
                FehlerGrund = "BAR Rate not available"
                Tabelle5.Range("I27").Value = FehlerGrund
                With Me.lbl_Quote_Info
                .Caption = "Due to the lack of a BAR Rate on the arrival date, not eligible for Quick Quote. Please contact your Revenue Manager."
                .ForeColor = RGB(255, 0, 0)
                .AutoSize = False
                .WordWrap = True
                End With
                Exit Sub
                
            Else
                
                sinSELLWish1 = sinSELL1 / (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value) + sinBPSplit
                sinSELLWalk1 = sinSELL1 / (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value) - sinSpielR_BME_(1) + sinBPSplit
            
            End If
        
        End If
        
        sinSellWishK1 = sinSELLWish1 '+ (sinSELLWish / 100 * Me.txt_Kommission)
        sinSellWalkK1 = sinSELLWalk1 '+ (sinSELLWalk / 100 * Me.txt_Kommission)
        
    'Tag 2
        
        If sinBME_Fix_(2) > sinBME_Min_(2) Then
            
            sinSELLWish2 = sinBME_Fix_(2)
            
        ElseIf sinBME_Min_(2) > sinBAR2 - sinBAR2 * sinBMEDiscount_(2) Then
        
            sinSELLWish2 = sinBME_Min_(2)
            
        Else
            
            If sinSELL2 = 0 Then
                
                Me.lbl_Rate.Caption = ""
                Me.lbl_Preisbereich.Caption = ""
                FehlerGrund = "BAR Rate not available"
                Tabelle5.Range("I27").Value = FehlerGrund
                With Me.lbl_Quote_Info
                .Caption = "Due to the lack of a BAR Rate on Day 2, not eligible for Quick Quote. Please contact your Revenue Manager."
                .ForeColor = RGB(255, 0, 0)
                .AutoSize = False
                .WordWrap = True
                End With
                Exit Sub
                
            Else
                
                sinSELLWish2 = sinSELL2 / (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value) + sinBPSplit
                sinSELLWalk2 = sinSELL2 / (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value) - sinSpielR_BME_(2) + sinBPSplit
            
            End If
                    
        End If
        
        sinSellWishK2 = sinSELLWish2 '+ (sinSELLWish / 100 * Me.txt_Kommission)
        sinSellWalkK2 = sinSELLWalk2 '+ (sinSELLWalk / 100 * Me.txt_Kommission)
        
    'Tag 3
        
        If sinBME_Fix_(3) > sinBME_Min_(3) Then
            
            sinSELLWish3 = sinBME_Fix_(3)
            
        ElseIf sinBME_Min_(3) > sinBAR3 - sinBAR3 * sinBMEDiscount_(3) Then
        
            sinSELLWish3 = sinBME_Min_(3)
            
        Else
            
            If sinSELL3 = 0 Then
                
                Me.lbl_Rate.Caption = ""
                Me.lbl_Preisbereich.Caption = ""
                FehlerGrund = "BAR Rate not available"
                Tabelle5.Range("I27").Value = FehlerGrund
                With Me.lbl_Quote_Info
                .Caption = "Due to the lack of a BAR Rate on Day 3, not eligible for Quick Quote. Please contact your Revenue Manager."
                .ForeColor = RGB(255, 0, 0)
                .AutoSize = False
                .WordWrap = True
                End With
                Exit Sub
                
            Else
                
                sinSELLWish3 = sinSELL3 / (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value) + sinBPSplit
                sinSELLWalk3 = sinSELL3 / (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value) - sinSpielR_BME_(3) + sinBPSplit
            
            End If
        
        End If
        
        sinSellWishK3 = sinSELLWish3 '+ (sinSELLWish / 100 * Me.txt_Kommission)
        sinSellWalkK3 = sinSELLWalk3 '+ (sinSELLWalk / 100 * Me.txt_Kommission)
            
    'Tag 4
        
        If sinBME_Fix_(4) > sinBME_Min_(4) Then
            
            sinSELLWish4 = sinBME_Fix_(4)
            
        ElseIf sinBME_Min_(4) > sinBAR4 - sinBAR4 * sinBMEDiscount_(4) Then
        
            sinSELLWish4 = sinBME_Min_(4)
            
        Else
            
            If sinSELL4 = 0 Then
                
                Me.lbl_Rate.Caption = ""
                Me.lbl_Preisbereich.Caption = ""
                FehlerGrund = "BAR Rate not available"
                Tabelle5.Range("I27").Value = FehlerGrund
                With Me.lbl_Quote_Info
                .Caption = "Due to the lack of a BAR Rate on Day 4, not eligible for Quick Quote. Please contact your Revenue Manager."
                .ForeColor = RGB(255, 0, 0)
                .AutoSize = False
                .WordWrap = True
                End With
                Exit Sub
                
            Else
                
                sinSELLWish4 = sinSELL4 / (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value) + sinBPSplit
                sinSELLWalk4 = sinSELL4 / (Me.txt_ZimmerAnzahl.Value + Me.txt_ZimmerAnzahl_dz.Value) - sinSpielR_BME_(4) + sinBPSplit
            
            End If
        
        End If
        
        sinSellWishK4 = sinSELLWish4 '+ (sinSELLWish / 100 * Me.txt_Kommission)
        sinSellWalkK4 = sinSELLWalk4 '+ (sinSELLWalk / 100 * Me.txt_Kommission)
            
        sinSellWishK = (sinSellWishK1 + sinSellWishK2 + sinSellWishK3 + sinSellWishK4) / intAnzTage
        sinSellWalkK = (sinSellWalkK1 + sinSellWalkK2 + sinSellWalkK3 + sinSellWalkK4) / intAnzTage
        
    End Select


'*************BERECHNUNG************


    endPreis = Application.WorksheetFunction.Max(sinSellWishK, sinSellWalkK)
    
    sin_Kat_avg = (sin_Kat_(1) + sin_Kat_(2) + sin_Kat_(3) + sin_Kat_(4)) / intAnzTage
    
    sinSpielR_BME = (sinSpielR_BME_(1) + sinSpielR_BME_(2) + sinSpielR_BME_(3) + sinSpielR_BME_(4)) / intAnzTage
    
    With Me.lbl_Preisbereich
    .Caption = "The price for the group:"
    End With
    
    With Me.lbl_Rate
    .Caption = Round(endPreis, 1) & " €"
    End With
    
    ' * Plenum Protect Prüfung *
    
    If Me.chk_plenum = True Then
        For i = 0 To intAnzTage - 1
            If Tabelle5.Range("AQ" & rowDiscount + i).Value = ChrW(&H2713) Then
                FehlerGrund = "There is no meeting room available. The price for the group would be: " & Round(Application.WorksheetFunction.Max(sinSellWishK, sinSellWalkK), 1) & " €"
                Tabelle5.Range("I27").Value = FehlerGrund
                With Me.lbl_Preisbereich
                .Caption = ""
                End With
                With Me.lbl_Rate
                .Caption = ""
                End With
                With Me.lbl_Quote_Info
                .Caption = "On the days of residence, there is no meeting room available." & Chr(10) & _
                "Please contact your Revenue Manager!"
                .ForeColor = RGB(255, 0, 0)
                .AutoSize = False
                .WordWrap = True
                End With
                Exit Sub
            End If
        Next i
    End If
    
    With Me.lbl_Quote_Info
    .Caption = "The prices include breakfast in single room." & Chr(10) & _
                "The category surcharge for the residence is " & sin_Kat_avg & " €." & Chr(10) & _
                "The surcharge for a double room is " & BGR_Config_v3.txt_bgr_dz & " €." & Chr(10) & _
                "The room for negotiation is " & sinSpielR_BME & " €."
    .ForeColor = RGB(0, 0, 0)
    .AutoSize = False
    .WordWrap = True
    End With
    
End Select

End Sub


'Schließen Button

'------------------Quick Quote schließen --------------------------------

Private Sub cmd_Schließen_Click()
Unload Quick_Quote_v3
End Sub

Private Sub cmd_conf_Schließen_Click()
Unload Quick_Quote_v3
Unload LGR_Config_v3
End Sub

'Basis Konfiguration

'------------------Konfiguration wird auf Reiter "mapping" gespeichert --------------------------------

Private Sub cmd_speichern_Click()

Application.Calculation = xlCalculationManual

'Tabelle5.Range("K2").Value = Me.txt_Disc_BRO.Value * 1
'Tabelle5.Range("K3").Value = Me.txt_Disc_BME.Value * 1
'Tabelle5.Range("K4").Value = Me.txt_SpielR_BRO.Value * 1
'Tabelle5.Range("K5").Value = Me.txt_SpielR_BME.Value * 1
'Tabelle5.Range("K6").Value = Me.txt_SpielR_LGR.Value * 1
Tabelle5.Range("K7").Value = Me.txt_BP_Split.Value * 1  'Spalte K
'Tabelle5.Range("K8").Value = Me.txt_max_room_bro.Value * 1
'Tabelle5.Range("K9").Value = Me.txt_OCC.Value * 1
'Tabelle5.Range("K10").Value = Me.txt_MLOS.Value * 1
'Tabelle5.Range("K11").Value = Me.chk_Total_Occ.Value
Tabelle5.Range("K12").Value = Me.chk_CXL_Policy.Value  'Spalte K
'Tabelle5.Range("K13").Value = Me.chk_Restriktion.Value
'Tabelle5.Range("K14").Value = Me.chk_MLOS.Value
Tabelle5.Range("K15").Value = Me.chk_Dyn_MLOS  'Spalte K
'Tabelle5.Range("K16").Value = Me.txt_bgr_dz.Value * 1
'Tabelle5.Range("K15").Value = Me.txt_max_room_bme.Value * 1
'Tabelle5.Range("K16").Value = Me.txt_max_room_lgr.Value * 1

Application.Calculation = xlCalculationAutomatic

MsgBox "The changes will be saved after the restart of Quick Quote."

End Sub

'Passwort aufrufen

Private Sub MultiPage1_Change()

If MultiPage1.Value = 1 Then

PW_v2.Show
End If

End Sub

Private Sub UserForm_Initialize()

'___________________________General_________________________

Me.Caption = "Quick Quote from " & Environ("UserName") & " on " & Date
Me.BackColor = RGB(255, 255, 255)
Me.StartUpPosition = 3


'___________________________MultiPage_________________________

With Me.MultiPage1
.BackColor = RGB(255, 255, 255)
.Font.Name = "Calibri"
.Pages(0).Caption = "Quick Quote"
.Pages(1).Caption = "Configuration"
'.Pages(0).BackColor = RGB(255, 255, 255)
'.Pages(1).BackColor = RGB(255, 255, 255)
End With

'___________________________CommandButton_________________________

'.........................Page1...........................

With Me.cmd_QuickQuote
'.BackColor = RGB(0, 205, 0)
.BackColor = RGB(240, 255, 255)
.Font.Name = "Calibri"
.Font.Size = 10
.Font.Bold = True
.Default = True
End With

With Me.cmd_Schließen
'.BackColor = RGB(255, 255, 255)
.BackColor = RGB(245, 255, 250)
.Caption = "Close"
End With

With Me.cmd_EMailErstellen
'.BackColor = RGB(0, 255, 255)
.BackColor = RGB(240, 255, 240)
.Caption = "Create E-Mail"
End With

'.........................Page2...........................

With Me.cmd_Speichern
.BackColor = RGB(255, 255, 255)
.Caption = "Save"
End With

With Me.cmd_conf_Schließen
.BackColor = RGB(255, 255, 255)
.Caption = "Close"
End With

With Me.cmd_LGR_Konfig
.BackColor = RGB(255, 255, 255)
.Caption = "LGR Config"
End With

With Me.cmd_BGR_Konfig
.BackColor = RGB(255, 255, 255)
.Caption = "BGR Config"
End With



'___________________________ComboBox_________________________

'.........................Page1...........................

With Me.cbo_Segment
.AddItem "BRO"
.AddItem "BME"
.AddItem "LGR"
.ListIndex = 0
.Style = fmStyleDropDownList
.TextAlign = fmTextAlignCenter
.BackColor = RGB(255, 255, 255)
.Font.Name = "Calibri"
.Font.Size = 10
End With

'___________________________Checkbox_________________________

'.........................Page1...........................

With Me.chk_plenum
.Value = True
.Enabled = False
.BackColor = RGB(255, 255, 255)
.Font.Name = "Calibri"
.Font.Size = 10
End With

'.........................Page2...........................

With Me.txt_BP_Split
.Value = Tabelle5.Range("K7").Value
.BackColor = RGB(255, 255, 255)
.Font.Name = "Calibri"
.Font.Size = 10
End With

With Me.chk_CXL_Policy
.Value = Tabelle5.Range("K12").Value
.BackColor = RGB(255, 255, 255)
.Font.Name = "Calibri"
.Font.Size = 10
End With

With Me.chk_Dyn_MLOS
.Value = Tabelle5.Range("K15").Value
.BackColor = RGB(255, 255, 255)
.Font.Name = "Calibri"
.Font.Size = 10
End With

With Me.chk_plenum
.Value = Tabelle5.Range("K17").Value
.BackColor = RGB(255, 255, 255)
.Font.Name = "Calibri"
.Font.Size = 10
End With

'___________________________Textbox_________________________

'.........................Page1...........................

With Me.txt_Anreise
.Value = Date + 45
.TextAlign = fmTextAlignCenter
.BackColor = RGB(255, 255, 255)
.Font.Name = "Calibri"
.Font.Size = 10
.SetFocus
.SelStart = 0
.SelLength = 10
End With

With Me.txt_Abreise
.Value = Date + 46
.TextAlign = fmTextAlignCenter
.BackColor = RGB(255, 255, 255)
.Font.Name = "Calibri"
.Font.Size = 10
End With

With Me.txt_ZimmerAnzahl
.Value = ""
.TextAlign = fmTextAlignCenter
.BackColor = RGB(255, 255, 255)
.Font.Name = "Calibri"
.Font.Size = 10
End With

With Me.txt_ZimmerAnzahl_dz
.Value = ""
.TextAlign = fmTextAlignCenter
.BackColor = RGB(255, 255, 255)
.Font.Name = "Calibri"
.Font.Size = 10
End With

With Me.txt_Kommission
.Value = "0"
.TextAlign = fmTextAlignCenter
.BackColor = RGB(255, 255, 255)
.Font.Name = "Calibri"
.Font.Size = 10
End With

'.........................Page2...........................


With Me.txt_BP_Split
.Value = Tabelle5.Range("K7")
.TextAlign = fmTextAlignCenter
.BackColor = RGB(255, 255, 255)
.Font.Name = "Calibri"
.Font.Size = 10
End With

'___________________________Label_________________________

'.........................Page1...........................

With Me.lbl_Hotelname
.BackColor = RGB(255, 255, 255)
.Font.Name = "Calibri"
.Font.Size = 12
.TextAlign = fmTextAlignLeft
.Caption = Tabelle18.Range("A6").Value
End With

With Me.lbl_Anreisedatum
.BackColor = RGB(255, 255, 255)
.Font.Name = "Calibri"
.Font.Size = 12
End With

With Me.lbl_Abreisedatum
.BackColor = RGB(255, 255, 255)
.Font.Name = "Calibri"
.Font.Size = 12
End With

With Me.lbl_ZimmerAnzahl
.BackColor = RGB(255, 255, 255)
.Font.Name = "Calibri"
.Font.Size = 12
End With

With Me.lbl_ZimmerAnzahl_dz
.BackColor = RGB(255, 255, 255)
.Font.Name = "Calibri"
.Font.Size = 12
End With

With Me.lbl_Segment
.BackColor = RGB(255, 255, 255)
.Font.Name = "Calibri"
.Font.Size = 12
End With

With Me.lbl_Kommission
.BackColor = RGB(255, 255, 255)
.Font.Name = "Calibri"
.Font.Size = 12
End With

With Me.lbl_Preisbereich
.BackColor = RGB(255, 255, 255)
.Font.Name = "Calibri"
.Font.Size = 12
.TextAlign = fmTextAlignRight
End With

With Me.lbl_Rate
.BackColor = RGB(255, 255, 255)
.Font.Name = "Calibri"
.Font.Size = 12
.TextAlign = fmTextAlignCenter
End With

With Me.lbl_Quote_Info
.BackColor = RGB(255, 255, 255)
.Font.Name = "Calibri"
.Font.Size = 12
.TextAlign = fmTextAlignLeft

End With

'.........................Page2...........................

With Me.lbl_BP_Split
.BackColor = RGB(255, 255, 255)
.Font.Name = "Calibri"
.Font.Size = 10
End With

End Sub


Sub LGR_2021()

Dim dateAnreise As String, dateAbreise As String
Dim dateAnreise2 As String, dateAnreise3 As String, dateAnreise4 As String
Dim i As Integer, j As Integer
Dim intZeileAn As Integer, intzeileAb As Integer
Dim intZeile2 As Integer, intZeile3 As Integer, intZeile4 As Integer
Dim intLGR_Row As Integer, intZeileAnLGR As Integer
Dim intZeile2LGR As Integer, intZeile3LGR As Integer, intZeile4LGR As Integer
Dim rngAnreise As Range, rngAbreise As Range
Dim rngTag2 As Range, rngTag3 As Range, rngTag4 As Range
Dim rngsaison1 As Range, rngsaison2 As Range, rngsaison3 As Range, rngsaison4 As Range
Dim LGR_Zeile As Integer, LGR_Zeile2 As Integer, LGR_Zeile3 As Integer, LGR_Zeile4 As Integer
Dim firstDate As Date, firstDate2 As Date, firstDate3 As Date, firstDate4 As Date
Dim intAnzTage As Integer
Dim Tag As String, Monat As String, Jahr As String, FehlerGrund As String

Dim sinBAR1 As Single, sinBAR2 As Single, sinBAR3 As Single, sinBAR4 As Single
Dim sinSELL1 As Single, sinSELL2 As Single, sinSELL3 As Single, sinSELL4 As Single
Dim sinSELLWalk As Single, sinSELLWish As Single

Dim cell         As Range

'Dim start2       As Date
'Dim finish2      As Date

Dim rngsaison    As Object
Dim intSaison    As Object
Dim intSaison1   As Long
Dim intSaison2   As Integer
Dim intSaison3   As Integer
Dim intSaison4   As Integer

Dim sinLGR1     As Single
Dim sinLGR2     As Single
Dim sinLGR3     As Single
Dim sinLGR4     As Single
Dim sinLGRRate  As Single

Dim sinEZZ      As Single
Dim sinEZZ2     As Single
Dim sinEZZ3     As Single
Dim sinEZZ4     As Single
Dim sinXbett    As Single
Dim sinHP       As Single
Dim sinVP       As Single
Dim sinHP2      As Single
Dim sinVP2      As Single
Dim sinHP3      As Single
Dim sinVP3      As Single
Dim sinHP4      As Single
Dim sinVP4      As Single
Dim sinGepaeck  As Single
Dim sinEZTeil   As Single
Dim endPreis    As Single
    
Dim start1 As Variant, finish1 As Variant, start2 As Variant, finish2 As Date

dateAnreise = Me.txt_Anreise
dateAbreise = Me.txt_Abreise
intAnzTage = CDate(dateAbreise) - CDate(dateAnreise)

Application.Calculation = xlCalculationManual
    
sinXbett = Tabelle5.Range("N3").Value
sinGepaeck = Tabelle5.Range("N4").Value

sinEZTeil = Me.txt_ZimmerAnzahl.Value * 1 / (Me.txt_ZimmerAnzahl.Value * 1 + Me.txt_ZimmerAnzahl_dz.Value * 1)

firstDate = CDate(Tabelle5.Range("N27").Value)
firstDate2 = CDate(Tabelle5.Range("N28").Value)
firstDate3 = CDate(Tabelle5.Range("N29").Value)
firstDate4 = CDate(Tabelle5.Range("N30").Value)

Select Case intAnzTage

'........................ / 1 Tag Aufenthalt / ...................


Case 1

Tabelle5.Range("N23").Value = CDate(Me.txt_Anreise.Value)
Tabelle5.Range("N24").Value = 0
Tabelle5.Range("N25").Value = 0
Tabelle5.Range("N26").Value = 0

Application.Calculate

If Tabelle5.Range("N27").Value = "" Then
    Me.lbl_Rate.Caption = ""
    Me.lbl_Preisbereich.Caption = ""
    With Me.lbl_Quote_Info
    .Caption = "There is no season entered in the system for the requested residence. Please contact your Revenue Manager."
    .ForeColor = RGB(255, 0, 0)
    .AutoSize = False
    .WordWrap = True
    End With
    bolexit = True
    Exit Sub
End If

intLGR_Row = Tabelle5.Range("O1048576").End(xlUp).row
firstDate = CDate(Tabelle5.Range("N27"))

With Tabelle5

'*******************************Anreisetag******************************

i = 1

start2 = CDate(.Range("N27").Value)
On Error Resume Next
finish2 = .Range("O1:O" & intLGR_Row).Find(CDate(start2)).Offset(0, 1).Value
If Err.Number <> 0 Then
finish2 = .Range("O1:O" & intLGR_Row).Find(CStr(start2)).Offset(0, 1).Value
End If
On Error GoTo 0

Set rngsaison = CreateObject("Scripting.Dictionary")

For Each cell In .Range("O2:O" & intLGR_Row)
    If cell.Value = firstDate And cell.Offset(0, 1).Value = finish2 Then
    rngsaison(i) = cell.row
    i = i + 1
    End If
Next

j = i - 1

For i = 1 To j

    If .Cells(rngsaison(i), .Rows(1).Find(what:=.Columns(1).Find(what:=Left(Format(CDate(dateAnreise), "ddd"), 2)).Offset(0, 1).Value).Column) = ChrW(&H2713) Then
        intSaison1 = rngsaison(i): Exit For
    ElseIf i = j Then
        FehlerGrund = .Columns(1).Find(what:=Left(Format(CDate(dateAnreise), "ddd"), 2)).Offset(0, 2).Value & " (arrival date) is excluded from Quick Quote."
        Me.lbl_Rate.Caption = ""
        Me.lbl_Preisbereich.Caption = ""
        With Me.lbl_Quote_Info
            .Caption = FehlerGrund & " Please contact your Revenue Manager."
            .ForeColor = RGB(255, 0, 0)
            .AutoSize = False
            .WordWrap = True
        End With
    bolexit = True
    Exit Sub
    End If
Next i

End With

'Ratenberechnung

Tabelle5.Range("N23").Value = CDate(Me.txt_Anreise.Value)
Tabelle5.Range("N24").Value = 0
Tabelle5.Range("N25").Value = 0
Tabelle5.Range("N26").Value = 0

With Tabelle5

    '*******************************Anreisetag******************************
    
    i = 1
    
    start2 = CDate(.Range("N27").Value)
    On Error Resume Next
    finish2 = .Range("O1:O" & intLGR_Row).Find(CDate(start2)).Offset(0, 1).Value
    If Err.Number <> 0 Then
    finish2 = .Range("O1:O" & intLGR_Row).Find(CStr(start2)).Offset(0, 1).Value
    End If
    On Error GoTo 0
    
    Set rngsaison = CreateObject("Scripting.Dictionary")
    
    For Each cell In .Range("O2:O" & intLGR_Row)
        If cell.Value = firstDate And cell.Offset(0, 1).Value = finish2 Then
        rngsaison(i) = cell.row
        i = i + 1
        End If
    Next
    
    j = i - 1
    
    For i = 1 To j
        If .Cells(rngsaison(i), .Rows(1).Find(what:=.Columns(1).Find(what:=Left(Format(CDate(dateAnreise), "ddd"), 2)).Offset(0, 1).Value).Column) = ChrW(&H2713) Then
            LGR_Zeile = rngsaison(i): Exit For
        End If
    Next i

End With

sinLGR1 = Tabelle5.Cells(LGR_Zeile, 17).Value
sinHP = Tabelle5.Cells(LGR_Zeile, 20).Value
sinVP = Tabelle5.Cells(LGR_Zeile, 21).Value

If sinEZTeil > 0.201 Then
sinEZZ = Tabelle5.Range("R" & LGR_Zeile).Value
Else
sinEZZ = Tabelle5.Range("S" & LGR_Zeile).Value
End If

sinLGRRate = Application.WorksheetFunction.Average(sinLGR1)

With Me.lbl_Preisbereich
.Caption = "Price per person and night in double room:"
End With

With Me.lbl_Rate
.Caption = Round(sinLGRRate, 1) & " €"
End With

With Me.lbl_Quote_Info
.Caption = "The prices include breakfast." & Chr(10) & _
            "Single room surcharge: " & sinEZZ & " €, " & "Extra bed: " & sinXbett & " €, " & "HB: " & Round(sinHP, 1) & " €," & "FB: " & Round(sinVP, 1) & " €," & Chr(10) & _
            "Luggage service: " & sinGepaeck & " €"
.ForeColor = RGB(0, 0, 0)
.AutoSize = False
.WordWrap = True
End With
'
''........................ / 2 Tage Aufenthalt / ...................
'

Case 2

Tabelle5.Range("N23").Value = CDate(Me.txt_Anreise.Value)
Tabelle5.Range("N24").Value = CDate(Me.txt_Anreise.Value) + 1
Tabelle5.Range("N25").Value = 0
Tabelle5.Range("N26").Value = 0

Application.Calculate

If Tabelle5.Range("N27").Value = "" Or Tabelle5.Range("N28").Value = "" Then
    Me.lbl_Rate.Caption = ""
    Me.lbl_Preisbereich.Caption = ""
    With Me.lbl_Quote_Info
    .Caption = "There is no season entered in the system for the requested residence. Please contact your Revenue Manager."
    .ForeColor = RGB(255, 0, 0)
    .AutoSize = False
    .WordWrap = True
    End With
    bolexit = True
    Exit Sub
End If

intLGR_Row = Tabelle5.Range("O1048576").End(xlUp).row
firstDate = CDate(Tabelle5.Range("N27"))
firstDate2 = CDate(Tabelle5.Range("N28"))

dateAnreise2 = CDate(dateAnreise) + 1

With Tabelle5

'*******************************Anreisetag******************************

i = 1

start2 = CDate(.Range("N27").Value)
On Error Resume Next
finish2 = .Range("O1:O" & intLGR_Row).Find(CDate(start2)).Offset(0, 1).Value
If Err.Number <> 0 Then
finish2 = .Range("O1:O" & intLGR_Row).Find(CStr(start2)).Offset(0, 1).Value
End If
On Error GoTo 0

Set rngsaison = CreateObject("Scripting.Dictionary")

For Each cell In .Range("O2:O" & intLGR_Row)
    If cell.Value = firstDate And cell.Offset(0, 1).Value = finish2 Then
    rngsaison(i) = cell.row
    i = i + 1
    End If
Next

j = i - 1

For i = 1 To j
    If .Cells(rngsaison(i), .Rows(1).Find(what:=.Columns(1).Find(what:=Left(Format(CDate(dateAnreise), "ddd"), 2)).Offset(0, 1).Value).Column) = ChrW(&H2713) Then
        intSaison1 = rngsaison(i): Exit For
    ElseIf i = j Then
        FehlerGrund = .Columns(1).Find(what:=Left(Format(CDate(dateAnreise), "ddd"), 2)).Offset(0, 2).Value & " (arrival date) is excluded from Quick Quote."
        Me.lbl_Rate.Caption = ""
        Me.lbl_Preisbereich.Caption = ""
        With Me.lbl_Quote_Info
            .Caption = FehlerGrund & " Please contact your Revenue Manager."
            .ForeColor = RGB(255, 0, 0)
            .AutoSize = False
            .WordWrap = True
        End With
    bolexit = True
    Exit Sub
    End If
Next i

'*******************************Tag 2******************************

i = 1

start2 = CDate(.Range("N28").Value)
On Error Resume Next
finish2 = .Range("O1:O" & intLGR_Row).Find(CDate(start2)).Offset(0, 1).Value
If Err.Number <> 0 Then
finish2 = .Range("O1:O" & intLGR_Row).Find(CStr(start2)).Offset(0, 1).Value
End If
On Error GoTo 0

Set rngsaison = CreateObject("Scripting.Dictionary")

For Each cell In .Range("O2:O" & intLGR_Row)
    If cell.Value = firstDate2 And cell.Offset(0, 1).Value = finish2 Then
    rngsaison(i) = cell.row
    i = i + 1
    End If
Next

j = i - 1

For i = 1 To j
    If .Cells(rngsaison(i), .Rows(1).Find(what:=.Columns(1).Find(what:=Left(Format(CDate(dateAnreise2), "ddd"), 2)).Offset(0, 1).Value).Column) = ChrW(&H2713) Then
        intSaison2 = rngsaison(i): Exit For
    ElseIf i = j Then
        FehlerGrund = .Columns(1).Find(what:=Left(Format(CDate(dateAnreise2), "ddd"), 2)).Offset(0, 2).Value & " (2nd day) is excluded from Quick Quote."
        Me.lbl_Rate.Caption = ""
        Me.lbl_Preisbereich.Caption = ""
        With Me.lbl_Quote_Info
            .Caption = FehlerGrund & " Please contact your Revenue Manager."
            .ForeColor = RGB(255, 0, 0)
            .AutoSize = False
            .WordWrap = True
        End With
    bolexit = True
    Exit Sub
    End If
Next i

End With

Tabelle5.Range("N23").Value = CDate(Me.txt_Anreise.Value)
Tabelle5.Range("N24").Value = CDate(Me.txt_Anreise.Value) + 1
Tabelle5.Range("N25").Value = 0
Tabelle5.Range("N26").Value = 0

'*******************************Ratenberechnung******************************

With Tabelle5

    '*******************************Anreisetag******************************
    
    i = 1
    
    start2 = CDate(.Range("N27").Value)
    On Error Resume Next
    finish2 = .Range("O1:O" & intLGR_Row).Find(CDate(start2)).Offset(0, 1).Value
    If Err.Number <> 0 Then
    finish2 = .Range("O1:O" & intLGR_Row).Find(CStr(start2)).Offset(0, 1).Value
    End If
    
    On Error GoTo 0
    
    Set rngsaison = CreateObject("Scripting.Dictionary")
    
    For Each cell In .Range("O2:O" & intLGR_Row)
        If cell.Value = firstDate And cell.Offset(0, 1).Value = finish2 Then
        rngsaison(i) = cell.row
        i = i + 1
        End If
    Next
    
    j = i - 1
    
    For i = 1 To j
        If .Cells(rngsaison(i), .Rows(1).Find(what:=.Columns(1).Find(what:=Left(Format(CDate(dateAnreise), "ddd"), 2)).Offset(0, 1).Value).Column) = ChrW(&H2713) Then
            LGR_Zeile = rngsaison(i): Exit For
        End If
    Next i
    
    '*******************************Tag 2******************************
    
    i = 1
    
    start2 = CDate(.Range("N28").Value)
    On Error Resume Next
    finish2 = .Range("O1:O" & intLGR_Row).Find(CDate(start2)).Offset(0, 1).Value
    If Err.Number <> 0 Then
    finish2 = .Range("O1:O" & intLGR_Row).Find(CStr(start2)).Offset(0, 1).Value
    End If
    
    On Error GoTo 0
    
    Set rngsaison = CreateObject("Scripting.Dictionary")
    
    For Each cell In .Range("O2:O" & intLGR_Row)
        If cell.Value = firstDate2 And cell.Offset(0, 1).Value = finish2 Then
        rngsaison(i) = cell.row
        i = i + 1
        End If
    Next
    
    j = i - 1
    
    For i = 1 To j
        If .Cells(rngsaison(i), .Rows(1).Find(what:=.Columns(1).Find(what:=Left(Format(CDate(dateAnreise2), "ddd"), 2)).Offset(0, 1).Value).Column) = ChrW(&H2713) Then
            LGR_Zeile2 = rngsaison(i): Exit For
        End If
    Next i

End With

'*****************************************************************

sinLGR1 = Tabelle5.Cells(LGR_Zeile, 17).Value
sinLGR2 = Tabelle5.Cells(LGR_Zeile2, 17).Value
sinHP = Tabelle5.Cells(LGR_Zeile, 20).Value
sinVP = Tabelle5.Cells(LGR_Zeile, 21).Value
sinHP2 = Tabelle5.Cells(LGR_Zeile2, 20).Value
sinVP2 = Tabelle5.Cells(LGR_Zeile2, 21).Value

sinLGRRate = Application.WorksheetFunction.Average(sinLGR1, sinLGR2)
sinHP = Application.WorksheetFunction.Average(sinHP, sinHP2)
sinVP = Application.WorksheetFunction.Average(sinVP, sinVP2)

If sinEZTeil > 0.201 Then
sinEZZ = Tabelle5.Range("R" & LGR_Zeile).Value
sinEZZ2 = Tabelle5.Range("R" & LGR_Zeile2).Value
sinEZZ = Application.WorksheetFunction.Average(sinEZZ, sinEZZ2)
Else
sinEZZ = Tabelle5.Range("S" & LGR_Zeile).Value
sinEZZ2 = Tabelle5.Range("S" & LGR_Zeile2).Value
sinEZZ = Application.WorksheetFunction.Average(sinEZZ, sinEZZ2)
End If

With Me.lbl_Preisbereich
.Caption = "Price per person and night in double room:"
End With

With Me.lbl_Rate
.Caption = Round(sinLGRRate, 1) & " €"
End With

With Me.lbl_Quote_Info
.Caption = "The prices include breakfast." & Chr(10) & _
            "Single room surcharge: " & sinEZZ & " €, " & "Extra bed: " & sinXbett & " €, " & "HB: " & Round(sinHP, 1) & " €," & "FB: " & Round(sinVP, 1) & " €," & Chr(10) & _
            "Luggage service: " & sinGepaeck & " €"
.ForeColor = RGB(0, 0, 0)
.AutoSize = False
.WordWrap = True
End With

'........................ / 3 Tage Aufenthalt / ...................


Case 3

Tabelle5.Range("N23").Value = CDate(Me.txt_Anreise.Value)
Tabelle5.Range("N24").Value = CDate(Me.txt_Anreise.Value) + 1
Tabelle5.Range("N25").Value = CDate(Me.txt_Anreise.Value) + 2
Tabelle5.Range("N26").Value = 0

Application.Calculate

If Tabelle5.Range("N27").Value = "" Or Tabelle5.Range("N28").Value = "" Or Tabelle5.Range("N29").Value = "" Then
    Me.lbl_Rate.Caption = ""
    Me.lbl_Preisbereich.Caption = ""
    With Me.lbl_Quote_Info
    .Caption = "There is no season entered in the system for the requested residence. Please contact your Revenue Manager."
    .ForeColor = RGB(255, 0, 0)
    .AutoSize = False
    .WordWrap = True
    End With
    bolexit = True
    Exit Sub
End If

intLGR_Row = Tabelle5.Range("O1048576").End(xlUp).row
firstDate = CDate(Tabelle5.Range("N27"))
firstDate2 = CDate(Tabelle5.Range("N28"))
firstDate3 = CDate(Tabelle5.Range("N29"))

dateAnreise2 = CDate(dateAnreise) + 1
dateAnreise3 = CDate(dateAnreise) + 2

With Tabelle5

'*******************************Anreisetag******************************

i = 1

start2 = CDate(.Range("N27").Value)
On Error Resume Next
finish2 = .Range("O1:O" & intLGR_Row).Find(CDate(start2)).Offset(0, 1).Value
If Err.Number <> 0 Then
finish2 = .Range("O1:O" & intLGR_Row).Find(CStr(start2)).Offset(0, 1).Value
End If
On Error GoTo 0

Set rngsaison = CreateObject("Scripting.Dictionary")

For Each cell In .Range("O2:O" & intLGR_Row)
    If cell.Value = firstDate And cell.Offset(0, 1).Value = finish2 Then
    rngsaison(i) = cell.row
    i = i + 1
    End If
Next

j = i - 1

For i = 1 To j
    If .Cells(rngsaison(i), .Rows(1).Find(what:=.Columns(1).Find(what:=Left(Format(CDate(dateAnreise), "ddd"), 2)).Offset(0, 1).Value).Column) = ChrW(&H2713) Then
        intSaison1 = rngsaison(i): Exit For
    ElseIf i = j Then
        FehlerGrund = .Columns(1).Find(what:=Left(Format(CDate(dateAnreise), "ddd"), 2)).Offset(0, 2).Value & " (arrival date) is excluded from Quick Quote."
        Me.lbl_Rate.Caption = ""
        Me.lbl_Preisbereich.Caption = ""
        With Me.lbl_Quote_Info
            .Caption = FehlerGrund & " Please contact your Revenue Manager."
            .ForeColor = RGB(255, 0, 0)
            .AutoSize = False
            .WordWrap = True
        End With
    bolexit = True
    Exit Sub
    End If
Next i


'*******************************Tag 2******************************

i = 1

start2 = CDate(.Range("N28").Value)
On Error Resume Next
finish2 = .Range("O1:O" & intLGR_Row).Find(CDate(start2)).Offset(0, 1).Value
If Err.Number <> 0 Then
finish2 = .Range("O1:O" & intLGR_Row).Find(CStr(start2)).Offset(0, 1).Value
End If
On Error GoTo 0

Set rngsaison = CreateObject("Scripting.Dictionary")

For Each cell In .Range("O2:O" & intLGR_Row)
    If cell.Value = firstDate2 And cell.Offset(0, 1).Value = finish2 Then
    rngsaison(i) = cell.row
    i = i + 1
    End If
Next

j = i - 1

For i = 1 To j
    If .Cells(rngsaison(i), .Rows(1).Find(what:=.Columns(1).Find(what:=Left(Format(CDate(dateAnreise2), "ddd"), 2)).Offset(0, 1).Value).Column) = ChrW(&H2713) Then
        intSaison2 = rngsaison(i): Exit For
    ElseIf i = j Then
        FehlerGrund = .Columns(1).Find(what:=Left(Format(CDate(dateAnreise2), "ddd"), 2)).Offset(0, 2).Value & " (2nd day) is excluded from Quick Quote."
        Me.lbl_Rate.Caption = ""
        Me.lbl_Preisbereich.Caption = ""
        With Me.lbl_Quote_Info
            .Caption = FehlerGrund & " Please contact your Revenue Manager."
            .ForeColor = RGB(255, 0, 0)
            .AutoSize = False
            .WordWrap = True
        End With
    bolexit = True
    Exit Sub
    End If
Next i

'*******************************Tag 3******************************

i = 1

start2 = CDate(.Range("N29").Value)
On Error Resume Next
finish2 = .Range("O1:O" & intLGR_Row).Find(CDate(start2)).Offset(0, 1).Value
If Err.Number <> 0 Then
finish2 = .Range("O1:O" & intLGR_Row).Find(CStr(start2)).Offset(0, 1).Value
End If
On Error GoTo 0

Set rngsaison = CreateObject("Scripting.Dictionary")

For Each cell In .Range("O2:O" & intLGR_Row)
    If cell.Value = firstDate3 And cell.Offset(0, 1).Value = finish2 Then
    rngsaison(i) = cell.row
    i = i + 1
    End If
Next

j = i - 1

For i = 1 To j
    If .Cells(rngsaison(i), .Rows(1).Find(what:=.Columns(1).Find(what:=Left(Format(CDate(dateAnreise3), "ddd"), 2)).Offset(0, 1).Value).Column) = ChrW(&H2713) Then
        intSaison3 = rngsaison(i): Exit For
    ElseIf i = j Then
        FehlerGrund = .Columns(1).Find(what:=Left(Format(CDate(dateAnreise3), "ddd"), 2)).Offset(0, 2).Value & " (3rd day) is excluded from Quick Quote."
        Me.lbl_Rate.Caption = ""
        Me.lbl_Preisbereich.Caption = ""
        With Me.lbl_Quote_Info
            .Caption = FehlerGrund & " Please contact your Revenue Manager."
            .ForeColor = RGB(255, 0, 0)
            .AutoSize = False
            .WordWrap = True
        End With
    bolexit = True
    Exit Sub
    End If
Next i

End With

Tabelle5.Range("N23").Value = CDate(Me.txt_Anreise.Value)
    Tabelle5.Range("N24").Value = CDate(Me.txt_Anreise.Value) + 1
    Tabelle5.Range("N25").Value = CDate(Me.txt_Anreise.Value) + 2
    Tabelle5.Range("N26").Value = 0
    
'*******************************Ratenberechnung******************************
    
With Tabelle5

    '*******************************Anreisetag******************************
    
    i = 1
    
    start2 = CDate(.Range("N27").Value)
    On Error Resume Next
    finish2 = .Range("O1:O" & intLGR_Row).Find(CDate(start2)).Offset(0, 1).Value
    If Err.Number <> 0 Then
    finish2 = .Range("O1:O" & intLGR_Row).Find(CStr(start2)).Offset(0, 1).Value
    End If
    
    On Error GoTo 0
    
    Set rngsaison = CreateObject("Scripting.Dictionary")
    
    For Each cell In .Range("O2:O" & intLGR_Row)
        If cell.Value = firstDate And cell.Offset(0, 1).Value = finish2 Then
        rngsaison(i) = cell.row
        i = i + 1
        End If
    Next
    
    j = i - 1
    
    For i = 1 To j
        If .Cells(rngsaison(i), .Rows(1).Find(what:=.Columns(1).Find(what:=Left(Format(CDate(dateAnreise), "ddd"), 2)).Offset(0, 1).Value).Column) = ChrW(&H2713) Then
            LGR_Zeile = rngsaison(i): Exit For
        End If
    Next i
    
    '*******************************Tag 2******************************
    
    i = 1
    
    start2 = CDate(.Range("N28").Value)
    On Error Resume Next
    finish2 = .Range("O1:O" & intLGR_Row).Find(CDate(start2)).Offset(0, 1).Value
    If Err.Number <> 0 Then
    finish2 = .Range("O1:O" & intLGR_Row).Find(CStr(start2)).Offset(0, 1).Value
    End If
    
    On Error GoTo 0
    
    Set rngsaison = CreateObject("Scripting.Dictionary")
    
    For Each cell In .Range("O2:O" & intLGR_Row)
        If cell.Value = firstDate2 And cell.Offset(0, 1).Value = finish2 Then
        rngsaison(i) = cell.row
        i = i + 1
        End If
    Next
    
    j = i - 1
    
    For i = 1 To j
        If .Cells(rngsaison(i), .Rows(1).Find(what:=.Columns(1).Find(what:=Left(Format(CDate(dateAnreise2), "ddd"), 2)).Offset(0, 1).Value).Column) = ChrW(&H2713) Then
            LGR_Zeile2 = rngsaison(i): Exit For
        End If
    Next i
    
    '*******************************Tag 3******************************
    
    i = 1
    
    start2 = CDate(.Range("N29").Value)
    On Error Resume Next
    finish2 = .Range("O1:O" & intLGR_Row).Find(CDate(start2)).Offset(0, 1).Value
    If Err.Number <> 0 Then
    finish2 = .Range("O1:O" & intLGR_Row).Find(CStr(start2)).Offset(0, 1).Value
    End If
    
    On Error GoTo 0
    
    Set rngsaison = CreateObject("Scripting.Dictionary")
    
    For Each cell In .Range("O2:O" & intLGR_Row)
        If cell.Value = firstDate3 And cell.Offset(0, 1).Value = finish2 Then
        rngsaison(i) = cell.row
        i = i + 1
        End If
    Next
    
    j = i - 1
    
    For i = 1 To j
        If .Cells(rngsaison(i), .Rows(1).Find(what:=.Columns(1).Find(what:=Left(Format(CDate(dateAnreise3), "ddd"), 2)).Offset(0, 1).Value).Column) = ChrW(&H2713) Then
            LGR_Zeile3 = rngsaison(i): Exit For
        End If
    Next i

End With

'*****************************************************************

sinLGR1 = Tabelle5.Cells(LGR_Zeile, 17).Value
sinLGR2 = Tabelle5.Cells(LGR_Zeile2, 17).Value
sinLGR3 = Tabelle5.Cells(LGR_Zeile3, 17).Value
sinHP = Tabelle5.Cells(LGR_Zeile, 20).Value
sinVP = Tabelle5.Cells(LGR_Zeile, 21).Value
sinHP2 = Tabelle5.Cells(LGR_Zeile2, 20).Value
sinVP2 = Tabelle5.Cells(LGR_Zeile2, 21).Value
sinHP3 = Tabelle5.Cells(LGR_Zeile3, 20).Value
sinVP3 = Tabelle5.Cells(LGR_Zeile3, 21).Value

sinLGRRate = Application.WorksheetFunction.Average(sinLGR1, sinLGR2, sinLGR3)
sinHP = Application.WorksheetFunction.Average(sinHP, sinHP2, sinHP3)
sinVP = Application.WorksheetFunction.Average(sinVP, sinVP2, sinVP3)

If sinEZTeil > 0.201 Then
sinEZZ = Tabelle5.Range("R" & LGR_Zeile).Value
sinEZZ2 = Tabelle5.Range("R" & LGR_Zeile2).Value
sinEZZ3 = Tabelle5.Range("R" & LGR_Zeile3).Value
sinEZZ = Application.WorksheetFunction.Average(sinEZZ, sinEZZ2, sinEZZ3)
Else
sinEZZ = Tabelle5.Range("S" & LGR_Zeile).Value
sinEZZ2 = Tabelle5.Range("S" & LGR_Zeile2).Value
sinEZZ3 = Tabelle5.Range("S" & LGR_Zeile3).Value
sinEZZ = Application.WorksheetFunction.Average(sinEZZ, sinEZZ2, sinEZZ3)
End If

With Me.lbl_Preisbereich
.Caption = "Price per person and night in double room:"
End With

With Me.lbl_Rate
.Caption = Round(sinLGRRate, 1) & " €"
End With

With Me.lbl_Quote_Info
.Caption = "The prices include breakfast." & Chr(10) & _
            "Single room surcharge: " & sinEZZ & " €, " & "Extra bed: " & sinXbett & " €, " & "HB: " & Round(sinHP, 1) & " €," & "FB: " & Round(sinVP, 1) & " €," & Chr(10) & _
            "Luggage service: " & sinGepaeck & " €"
.ForeColor = RGB(0, 0, 0)
.AutoSize = False
.WordWrap = True
End With

''........................ / 4 Tage Aufenthalt / ...................
'
Case 4

Tabelle5.Range("N23").Value = CDate(Me.txt_Anreise.Value)
Tabelle5.Range("N24").Value = CDate(Me.txt_Anreise.Value) + 1
Tabelle5.Range("N25").Value = CDate(Me.txt_Anreise.Value) + 2
Tabelle5.Range("N26").Value = CDate(Me.txt_Anreise.Value) + 3

Application.Calculate

If Tabelle5.Range("N27").Value = "" Or Tabelle5.Range("N28").Value = "" Or Tabelle5.Range("N29").Value = "" Or Tabelle5.Range("N30").Value = "" Then
    Me.lbl_Rate.Caption = ""
    Me.lbl_Preisbereich.Caption = ""
    With Me.lbl_Quote_Info
    .Caption = "There is no season entered in the system for the requested residence. Please contact your Revenue Manager."
    .ForeColor = RGB(255, 0, 0)
    .AutoSize = False
    .WordWrap = True
    End With
    bolexit = True
    Exit Sub
End If

intLGR_Row = Tabelle5.Range("O1048576").End(xlUp).row
firstDate = CDate(Tabelle5.Range("N27"))
firstDate2 = CDate(Tabelle5.Range("N28"))
firstDate3 = CDate(Tabelle5.Range("N29"))
firstDate4 = CDate(Tabelle5.Range("N30"))

dateAnreise2 = CDate(dateAnreise) + 1
dateAnreise3 = CDate(dateAnreise) + 2
dateAnreise4 = CDate(dateAnreise) + 3

With Tabelle5

'*******************************Anreisetag******************************

i = 1

start2 = CDate(.Range("N27").Value)
On Error Resume Next
finish2 = .Range("O1:O" & intLGR_Row).Find(CDate(start2)).Offset(0, 1).Value
If Err.Number <> 0 Then
finish2 = .Range("O1:O" & intLGR_Row).Find(CStr(start2)).Offset(0, 1).Value
End If
On Error GoTo 0

Set rngsaison = CreateObject("Scripting.Dictionary")

For Each cell In .Range("O2:O" & intLGR_Row)
    If cell.Value = firstDate And cell.Offset(0, 1).Value = finish2 Then
    rngsaison(i) = cell.row
    i = i + 1
    End If
Next

j = i - 1

For i = 1 To j
    If .Cells(rngsaison(i), .Rows(1).Find(what:=.Columns(1).Find(what:=Left(Format(CDate(dateAnreise), "ddd"), 2)).Offset(0, 1).Value).Column) = ChrW(&H2713) Then
        intSaison1 = rngsaison(i): Exit For
    ElseIf i = j Then
        FehlerGrund = .Columns(1).Find(what:=Left(Format(CDate(dateAnreise), "ddd"), 2)).Offset(0, 2).Value & " (arrival date) is excluded from Quick Quote."
        Me.lbl_Rate.Caption = ""
        Me.lbl_Preisbereich.Caption = ""
        With Me.lbl_Quote_Info
            .Caption = FehlerGrund & " Please contact your Revenue Manager."
            .ForeColor = RGB(255, 0, 0)
            .AutoSize = False
            .WordWrap = True
        End With
    bolexit = True
    Exit Sub
    End If
Next i


'*******************************Tag 2******************************

i = 1

start2 = CDate(.Range("N28").Value)
On Error Resume Next
finish2 = .Range("O1:O" & intLGR_Row).Find(CDate(start2)).Offset(0, 1).Value
If Err.Number <> 0 Then
finish2 = .Range("O1:O" & intLGR_Row).Find(CStr(start2)).Offset(0, 1).Value
End If
On Error GoTo 0

Set rngsaison = CreateObject("Scripting.Dictionary")

For Each cell In .Range("O2:O" & intLGR_Row)
    If cell.Value = firstDate2 And cell.Offset(0, 1).Value = finish2 Then
    rngsaison(i) = cell.row
    i = i + 1
    End If
Next

j = i - 1

For i = 1 To j
    If .Cells(rngsaison(i), .Rows(1).Find(what:=.Columns(1).Find(what:=Left(Format(CDate(dateAnreise2), "ddd"), 2)).Offset(0, 1).Value).Column) = ChrW(&H2713) Then
        intSaison2 = rngsaison(i): Exit For
    ElseIf i = j Then
        Me.lbl_Rate.Caption = ""
        Me.lbl_Preisbereich.Caption = ""
        FehlerGrund = .Columns(1).Find(what:=Left(Format(CDate(dateAnreise2), "ddd"), 2)).Offset(0, 2).Value & " (2nd day) is excluded from Quick Quote."
        With Me.lbl_Quote_Info
            .Caption = FehlerGrund & " Please contact your Revenue Manager."
            .ForeColor = RGB(255, 0, 0)
            .AutoSize = False
            .WordWrap = True
        End With
    bolexit = True
    Exit Sub
    End If
Next i

'*******************************Tag 3******************************

i = 1

start2 = CDate(.Range("N29").Value)
On Error Resume Next
finish2 = .Range("O1:O" & intLGR_Row).Find(CDate(start2)).Offset(0, 1).Value
If Err.Number <> 0 Then
finish2 = .Range("O1:O" & intLGR_Row).Find(CStr(start2)).Offset(0, 1).Value
End If
On Error GoTo 0

Set rngsaison = CreateObject("Scripting.Dictionary")

For Each cell In .Range("O2:O" & intLGR_Row)
    If cell.Value = firstDate3 And cell.Offset(0, 1).Value = finish2 Then
    rngsaison(i) = cell.row
    i = i + 1
    End If
Next

j = i - 1

For i = 1 To j
    If .Cells(rngsaison(i), .Rows(1).Find(what:=.Columns(1).Find(what:=Left(Format(CDate(dateAnreise3), "ddd"), 2)).Offset(0, 1).Value).Column) = ChrW(&H2713) Then
        intSaison3 = rngsaison(i): Exit For
    ElseIf i = j Then
        FehlerGrund = .Columns(1).Find(what:=Left(Format(CDate(dateAnreise3), "ddd"), 2)).Offset(0, 2).Value & " (3rd day) is excluded from Quick Quote."
        Me.lbl_Rate.Caption = ""
        Me.lbl_Preisbereich.Caption = ""
        With Me.lbl_Quote_Info
            .Caption = FehlerGrund & " Please contact your Revenue Manager."
            .ForeColor = RGB(255, 0, 0)
            .AutoSize = False
            .WordWrap = True
        End With
    bolexit = True
    Exit Sub
    End If
Next i

'*******************************Tag 4******************************

i = 1

start2 = CDate(.Range("N30").Value)
On Error Resume Next
finish2 = .Range("O1:O" & intLGR_Row).Find(CDate(start2)).Offset(0, 1).Value
If Err.Number <> 0 Then
finish2 = .Range("O1:O" & intLGR_Row).Find(CStr(start2)).Offset(0, 1).Value
End If
On Error GoTo 0

Set rngsaison = CreateObject("Scripting.Dictionary")

For Each cell In .Range("O2:O" & intLGR_Row)
    If cell.Value = firstDate4 And cell.Offset(0, 1).Value = finish2 Then
    rngsaison(i) = cell.row
    i = i + 1
    End If
Next

j = i - 1

For i = 1 To j
    If .Cells(rngsaison(i), .Rows(1).Find(what:=.Columns(1).Find(what:=Left(Format(CDate(dateAnreise4), "ddd"), 2)).Offset(0, 1).Value).Column) = ChrW(&H2713) Then
        intSaison4 = rngsaison(i): Exit For
    ElseIf i = j Then
        FehlerGrund = .Columns(1).Find(what:=Left(Format(CDate(dateAnreise4), "ddd"), 2)).Offset(0, 2).Value & " (4th day) is excluded from Quick Quote."
        Me.lbl_Rate.Caption = ""
        Me.lbl_Preisbereich.Caption = ""
        With Me.lbl_Quote_Info
            .Caption = FehlerGrund & " Please contact your Revenue Manager."
            .ForeColor = RGB(255, 0, 0)
            .AutoSize = False
            .WordWrap = True
        End With
    bolexit = True
    Exit Sub
    End If
Next i

'*******************************Ratenberechnung******************************

Tabelle5.Range("N23").Value = CDate(Me.txt_Anreise.Value)
Tabelle5.Range("N24").Value = CDate(Me.txt_Anreise.Value) + 1
Tabelle5.Range("N25").Value = CDate(Me.txt_Anreise.Value) + 2
Tabelle5.Range("N26").Value = CDate(Me.txt_Anreise.Value) + 3

With Tabelle5

    '*******************************Anreisetag******************************
    
    i = 1
    
    start2 = CDate(.Range("N27").Value)
    On Error Resume Next
    finish2 = .Range("O1:O" & intLGR_Row).Find(CDate(start2)).Offset(0, 1).Value
    If Err.Number <> 0 Then
    finish2 = .Range("O1:O" & intLGR_Row).Find(CStr(start2)).Offset(0, 1).Value
    End If
    
    On Error GoTo 0
    
    Set rngsaison = CreateObject("Scripting.Dictionary")
    
    For Each cell In .Range("O2:O" & intLGR_Row)
        If cell.Value = firstDate And cell.Offset(0, 1).Value = finish2 Then
        rngsaison(i) = cell.row
        i = i + 1
        End If
    Next
    
    j = i - 1
    
    For i = 1 To j
        If .Cells(rngsaison(i), .Rows(1).Find(what:=.Columns(1).Find(what:=Left(Format(CDate(dateAnreise), "ddd"), 2)).Offset(0, 1).Value).Column) = ChrW(&H2713) Then
            LGR_Zeile = rngsaison(i): Exit For
        End If
    Next i
    
    '*******************************Tag 2******************************
    
    i = 1
    
    start2 = CDate(.Range("N28").Value)
    On Error Resume Next
    finish2 = .Range("O1:O" & intLGR_Row).Find(CDate(start2)).Offset(0, 1).Value
    If Err.Number <> 0 Then
    finish2 = .Range("O1:O" & intLGR_Row).Find(CStr(start2)).Offset(0, 1).Value
    End If
    
    On Error GoTo 0
    
    Set rngsaison = CreateObject("Scripting.Dictionary")
    
    For Each cell In .Range("O2:O" & intLGR_Row)
        If cell.Value = firstDate2 And cell.Offset(0, 1).Value = finish2 Then
        rngsaison(i) = cell.row
        i = i + 1
        End If
    Next
    
    j = i - 1
    
    For i = 1 To j
        If .Cells(rngsaison(i), .Rows(1).Find(what:=.Columns(1).Find(what:=Left(Format(CDate(dateAnreise2), "ddd"), 2)).Offset(0, 1).Value).Column) = ChrW(&H2713) Then
            LGR_Zeile2 = rngsaison(i): Exit For
        End If
    Next i
    
    '*******************************Tag 3******************************
    
    i = 1
    
    start2 = CDate(.Range("N29").Value)
    On Error Resume Next
    finish2 = .Range("O1:O" & intLGR_Row).Find(CDate(start2)).Offset(0, 1).Value
    If Err.Number <> 0 Then
    finish2 = .Range("O1:O" & intLGR_Row).Find(CStr(start2)).Offset(0, 1).Value
    End If
    
    On Error GoTo 0
    
    Set rngsaison = CreateObject("Scripting.Dictionary")
    
    For Each cell In .Range("O2:O" & intLGR_Row)
        If cell.Value = firstDate3 And cell.Offset(0, 1).Value = finish2 Then
        rngsaison(i) = cell.row
        i = i + 1
        End If
    Next
    
    j = i - 1
    
    For i = 1 To j
        If .Cells(rngsaison(i), .Rows(1).Find(what:=.Columns(1).Find(what:=Left(Format(CDate(dateAnreise3), "ddd"), 2)).Offset(0, 1).Value).Column) = ChrW(&H2713) Then
            LGR_Zeile3 = rngsaison(i): Exit For
        End If
    Next i
    
    '*******************************Tag 4******************************
    
    i = 1
    
    start2 = CDate(.Range("N30").Value)
    On Error Resume Next
    finish2 = .Range("O1:O" & intLGR_Row).Find(CDate(start2)).Offset(0, 1).Value
    If Err.Number <> 0 Then
    finish2 = .Range("O1:O" & intLGR_Row).Find(CStr(start2)).Offset(0, 1).Value
    End If
    
    On Error GoTo 0
    
    Set rngsaison = CreateObject("Scripting.Dictionary")
    
    For Each cell In .Range("O2:O" & intLGR_Row)
        If cell.Value = firstDate4 And cell.Offset(0, 1).Value = finish2 Then
        rngsaison(i) = cell.row
        i = i + 1
        End If
    Next
    
    j = i - 1
    
    For i = 1 To j
        If .Cells(rngsaison(i), .Rows(1).Find(what:=.Columns(1).Find(what:=Left(Format(CDate(dateAnreise4), "ddd"), 2)).Offset(0, 1).Value).Column) = ChrW(&H2713) Then
            LGR_Zeile4 = rngsaison(i): Exit For
        End If
    Next i

End With

'*****************************************************************

sinLGR1 = Tabelle5.Cells(LGR_Zeile, 17).Value   'Spalte Q
sinLGR2 = Tabelle5.Cells(LGR_Zeile2, 17).Value   'Spalte Q
sinLGR3 = Tabelle5.Cells(LGR_Zeile3, 17).Value   'Spalte Q
sinLGR4 = Tabelle5.Cells(LGR_Zeile4, 17).Value   'Spalte Q
sinHP = Tabelle5.Cells(LGR_Zeile, 20).Value   'Spalte T
sinVP = Tabelle5.Cells(LGR_Zeile, 21).Value   'Spalte U
sinHP2 = Tabelle5.Cells(LGR_Zeile2, 20).Value   'Spalte T
sinVP2 = Tabelle5.Cells(LGR_Zeile2, 21).Value   'Spalte U
sinHP3 = Tabelle5.Cells(LGR_Zeile3, 20).Value   'Spalte T
sinVP3 = Tabelle5.Cells(LGR_Zeile3, 21).Value   'Spalte U
sinHP4 = Tabelle5.Cells(LGR_Zeile4, 20).Value   'Spalte T
sinVP4 = Tabelle5.Cells(LGR_Zeile4, 21).Value   'Spalte U

sinLGRRate = Application.WorksheetFunction.Average(sinLGR1, sinLGR2, sinLGR3, sinLGR4)
sinHP = Application.WorksheetFunction.Average(sinHP, sinHP2, sinHP3, sinHP4)
sinVP = Application.WorksheetFunction.Average(sinVP, sinVP2, sinVP3, sinVP4)

If sinEZTeil > 0.201 Then
sinEZZ = Tabelle5.Range("R" & LGR_Zeile).Value   'Spalte R
sinEZZ2 = Tabelle5.Range("R" & LGR_Zeile2).Value   'Spalte R
sinEZZ3 = Tabelle5.Range("R" & LGR_Zeile3).Value   'Spalte R
sinEZZ4 = Tabelle5.Range("R" & LGR_Zeile4).Value   'Spalte R
sinEZZ = Application.WorksheetFunction.Average(sinEZZ, sinEZZ2, sinEZZ3)
Else
sinEZZ = Tabelle5.Range("S" & LGR_Zeile).Value   'Spalte S
sinEZZ2 = Tabelle5.Range("S" & LGR_Zeile2).Value   'Spalte S
sinEZZ3 = Tabelle5.Range("S" & LGR_Zeile3).Value   'Spalte S
sinEZZ4 = Tabelle5.Range("S" & LGR_Zeile4).Value   'Spalte S
sinEZZ = Application.WorksheetFunction.Average(sinEZZ, sinEZZ2, sinEZZ3)
End If

If Tabelle5.Range("N4").Value > sinLGRRate Then 'Zelle N4
    endPreis = Tabelle5.Range("N4").Value   'Zelle N4
Else
    endPreis = sinLGRRate
End If

With Me.lbl_Preisbereich
.Caption = "Price per person and night in double room:"
End With

With Me.lbl_Rate
.Caption = Round(endPreis, 1) & " €"
End With

With Me.lbl_Quote_Info
.Caption = "The prices include breakfast." & Chr(10) & _
            "Single room surcharge: " & sinEZZ & " €, " & "Extra bed: " & sinXbett & " €, " & "HB: " & Round(sinHP, 1) & " €," & "FB: " & Round(sinVP, 1) & " €," & Chr(10) & _
            "Luggage service: " & sinGepaeck & " €"
.ForeColor = RGB(0, 0, 0)
.AutoSize = False
.WordWrap = True
End With

End With

End Select
'**************************************************************

Application.Calculation = xlCalculationAutomatic
'
End Sub

Private Sub cbo_Segment_Change()

If cbo_Segment.Value = "BME" Then
chk_plenum.Enabled = True
chk_plenum = True
Else
chk_plenum.Enabled = False
'chk_plenum.Locked = True
End If

End Sub

Private Sub cmd_EMailErstellen_Click()

Dim objApp As Object
Dim objMailItm As Object
Dim intCounter As Integer
Dim varDest As Variant
Dim strDest As String
Dim xInspect As Object
Dim pageEditor As Object
Dim FehlerGrund As String
Dim An As String
Dim CC As String
Dim intVerteiler_row As Long
Dim intCC_row As Long
Dim cell As Range
   
'Vorbereitung der E-Mail
Set objApp = CreateObject("Outlook.Application")
Set objMailItm = objApp.createitem(0)
   
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual
   
Tabelle5.Range("J18").Value = txt_Anreise.Text
Tabelle5.Range("J19").Value = txt_Abreise.Text
Tabelle5.Range("J20").Value = txt_ZimmerAnzahl.Text
Tabelle5.Range("J21").Value = txt_ZimmerAnzahl_dz.Text
Tabelle5.Range("J22").Value = cbo_Segment.Text
Tabelle5.Range("J23").Value = txt_Kommission.Text
FehlerGrund = Tabelle5.Range("I27").Value

If FehlerGrund = "" Then
    FehlerGrund = "Quotation"
End If

Tabelle5.Range("I24").Value = lbl_Preisbereich
Tabelle5.Range("J24").Value = lbl_Rate
Tabelle5.Range("I26").Value = lbl_Quote_Info
With Tabelle5.Range("I26")
 .WrapText = False
End With
  
Application.Calculation = xlCalculationAutomatic
'Unload Quick_Quote_v3
    
'intVerteiler_row = Tabelle5.Range("F1048576").End(xlUp).row
'
'If intVerteiler_row = 17 Then
'    intVerteiler_row = 18
'End If
'
'intCC_row = Tabelle5.Range("G1048576").End(xlUp).row
'
'If intCC_row = 17 Then
'    intCC_row = 18
'End If
'
'For Each cell In Tabelle5.Range("F18:F" & intVerteiler_row)
'    An = An & ";" & cell.Value
'Next cell
'
'For Each cell In Tabelle5.Range("G18:G" & intCC_row)
'    CC = CC & ";" & cell.Value
'Next cell
'
'Erstellung der E-Mail

With objMailItm
    On Error Resume Next
    .To = An
    .CC = CC
    .Subject = "Quick Quote " & Tabelle18.Range("A6").Value & " - " & FehlerGrund
    If FehlerGrund = "Quotation" Then
    .body = vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Please find below the Quick Quote results:" & vbCrLf & vbCrLf
    Else
    .body = vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Please find below the Quick Quote error message:" & vbCrLf & vbCrLf
    End If
    Set xInspect = objMailItm.GetInspector
    Set pageEditor = xInspect.WordEditor
    
    Tabelle5.Range("I18:J26").Copy
    Application.Wait (Now + TimeValue("0:00:04"))
    pageEditor.Application.Selection.Start = Len(.body)
    pageEditor.Application.Selection.End = pageEditor.Application.Selection.Start
    '.display
    pageEditor.Application.Selection.PasteAndFormat Type:=20
     
    .display
    Set pageEditor = Nothing
    Set xInspect = Nothing
End With

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.CutCopyMode = False
   
'Beenden

Set objMailItm = Nothing
Set objApp = Nothing

End Sub

Private Sub cmd_BGR_Konfig_Click()

On Error GoTo Error

BGR_Config_v3.Show vbModeless

Error: Exit Sub

End Sub

Private Sub CMD_LGR_Konfig_Click()

LGR_Config_v3.Show vbModeless

End Sub

Function DateiInBearbeitung(strDatei As String) As Boolean

On Error Resume Next
Open strDatei For Binary Access Read Lock Read As #1
Close #1
If Err.Number <> 0 Then
DateiInBearbeitung = True
Err.Clear
End If
End Function

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

Function CellContentCanBeInterpretedAsADate(cell As Range) As Boolean

Dim d As Date

On Error Resume Next

d = CDate(cell.Value)

If Err.Number <> 0 Then
    CellContentCanBeInterpretedAsADate = False
Else
    CellContentCanBeInterpretedAsADate = True
End If

On Error GoTo 0
    
End Function



