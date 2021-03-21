VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GÝRÝS 
   Caption         =   "Tedarik Zinciri Yöntemi"
   ClientHeight    =   3855
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8805.001
   OleObjectBlob   =   "GÝRÝS.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GÝRÝS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
Dim parola As String
Dim ID As String
        ID = GÝRÝS.TextBox1.Value
        parola = GÝRÝS.TextBox2.Value
     
        If ID = "admin@irse.deu.com.tr" And parola = "123456" Then
            'Worksheets("Tedarik Zinciri Yönetimi").Unprotect
            'Sheets("Amaç F. ve Kýsýtlar").Visible = True
            'Sheets("Karar Destek Sistemi").Visible = True
            'Sheets("KARAR DESTEK").Select
            GÝRÝS.Hide
            ANASAYFA.Show
            GÝRÝS.TextBox1.Value = ""
            GÝRÝS.TextBox2.Value = ""

        Else
            Call MsgBox("Kullanýcý Adý veya Parolanýz Hatalýdýr.Lütfen Tekrar Deneyiniz.", , "Tedarik Zinciri Yöntemi")
            'Sheets("Amaç F. ve Kýsýtlar").Visible = False
            'Sheets("Karar Destek Sistemi").Visible = False
            'Worksheets("Tedarik Zinciri Yönetimi").Protect
            GÝRÝS.TextBox1.Value = "admin@irse.deu.com.tr"
            GÝRÝS.TextBox2.Value = ""

        End If

End Sub

Private Sub CommandButton2_Click()
            Sheets("TEDARÝK ZÝNCÝRÝ YÖNETÝMÝ").Select
            GÝRÝS.Hide

End Sub


