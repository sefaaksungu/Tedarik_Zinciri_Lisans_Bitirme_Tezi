VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} karardestekerisim 
   Caption         =   "Tedarik Zinciri Yöntemi"
   ClientHeight    =   3840
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8910.001
   OleObjectBlob   =   "karardestekerisim.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "karardestekerisim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton3_Click()
Dim parola As String
Dim ID As String
        ID = karardestekerisim.TextBox3.Value
        parola = karardestekerisim.TextBox4.Value
     
        If ID = "admin@irse.deu.com.tr" And parola = "123456" Then
            'Worksheets("Tedarik Zinciri Yönetimi").Unprotect
            'Sheets("Amaç F. ve Kýsýtlar").Visible = True
            'Sheets("Karar Destek Sistemi").Visible = True
            Sheets("KARAR DESTEK").Select
            karardestekerisim.Hide
            ANASAYFA.Hide
            karardestekerisim.TextBox3.Value = ""
            karardestekerisim.TextBox4.Value = ""

        Else
            Call MsgBox("Kullanýcý Adý veya Parolanýz Hatalýdýr.Lütfen Tekrar Deneyiniz.", , "Tedarik Zinciri Yöntemi")
            'Sheets("Amaç F. ve Kýsýtlar").Visible = False
            'Sheets("Karar Destek Sistemi").Visible = False
            'Worksheets("Tedarik Zinciri Yönetimi").Protect
            karardestekerisim.TextBox3.Value = "admin@irse.deu.com.tr"
            karardestekerisim.TextBox4.Value = ""

        End If

End Sub

Private Sub CommandButton4_Click()
            karardestekerisim.Hide
            Sheets("TEDARÝK ZÝNCÝRÝ YÖNETÝMÝ").Select
End Sub

Private Sub TextBox3_Change()

End Sub
