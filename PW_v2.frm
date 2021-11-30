VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PW_v2 
   Caption         =   "Password"
   ClientHeight    =   3180
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   4704
   OleObjectBlob   =   "PW_v2.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "PW_v2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Ok_Click()
If Me.txt_Passwort.Value = "Hhotels18" Then
Unload Me
Else
MsgBox "The password is not correct." ' & Chr(10) & _
        '"Quick Quote will be cancelled." & Chr(10) & _
        '"Should you have more questions, please contact your Revenue Manager."
'Unload Me
'Unload Quick_Quote_v3
End If

End Sub

Private Sub cmd_PWCancel_Click()
Unload Me
'Quick_Quote_v3.MultiPage1.Value = 0: Quick_Quote_v3.Show vbModeless
Unload Quick_Quote_v3
Quick_Quote_v3.Show vbModeless
End Sub

Private Sub UserForm_Initialize()
Me.Caption = "Password Configuration"
Me.BackColor = RGB(255, 255, 255)
Me.StartUpPosition = 3

With Me.txt_Passwort
.TextAlign = fmTextAlignLeft
.BackColor = RGB(255, 255, 255)
.Font.Name = "Calibri"
.Font.Size = 12
.PasswordChar = "X"
.SetFocus
.ControlTipText = "Please enter a password."
.Font.Bold = True
.Height = 22
End With

With Me.lbl_Passwort_Eingabe
.Caption = "Enter password"
.Font.Size = 14
.Font.Bold = True
.Font.Name = "Calibri"
.TextAlign = fmTextAlignLeft
.Height = 18
End With

With Me.cmd_Ok
.BackColor = RGB(255, 255, 255)
.Font.Name = "Calibri"
.Font.Size = 10
.Caption = "OK"
.Default = True
End With

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If CloseMode = 0 Then
Cancel = True
End If
End Sub
