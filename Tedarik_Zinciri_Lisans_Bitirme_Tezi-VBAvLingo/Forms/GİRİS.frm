VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} G�R�S 
   Caption         =   "Tedarik Zinciri Y�ntemi"
   ClientHeight    =   3855
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8805.001
   OleObjectBlob   =   "G�R�S.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "G�R�S"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
Dim parola As String
Dim ID As String
        ID = G�R�S.TextBox1.Value
        parola = G�R�S.TextBox2.Value
     
        If ID = "admin@irse.deu.com.tr" And parola = "123456" Then
            'Worksheets("Tedarik Zinciri Y�netimi").Unprotect
            'Sheets("Ama� F. ve K�s�tlar").Visible = True
            'Sheets("Karar Destek Sistemi").Visible = True
            'Sheets("KARAR DESTEK").Select
            G�R�S.Hide
            ANASAYFA.Show
            G�R�S.TextBox1.Value = ""
            G�R�S.TextBox2.Value = ""

        Else
            Call MsgBox("Kullan�c� Ad� veya Parolan�z Hatal�d�r.L�tfen Tekrar Deneyiniz.", , "Tedarik Zinciri Y�ntemi")
            'Sheets("Ama� F. ve K�s�tlar").Visible = False
            'Sheets("Karar Destek Sistemi").Visible = False
            'Worksheets("Tedarik Zinciri Y�netimi").Protect
            G�R�S.TextBox1.Value = "admin@irse.deu.com.tr"
            G�R�S.TextBox2.Value = ""

        End If

End Sub

Private Sub CommandButton2_Click()
            Sheets("TEDAR�K Z�NC�R� Y�NET�M�").Select
            G�R�S.Hide

End Sub


