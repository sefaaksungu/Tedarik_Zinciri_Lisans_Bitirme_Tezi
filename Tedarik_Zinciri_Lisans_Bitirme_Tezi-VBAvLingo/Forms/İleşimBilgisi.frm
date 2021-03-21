VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ÝleþimBilgisi 
   Caption         =   "Ýletiþim Bilgileri"
   ClientHeight    =   6780
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9150.001
   OleObjectBlob   =   "ÝleþimBilgisi.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "ÝleþimBilgisi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub GeriButonu_Click()
ÝleþimBilgisi.Hide
End Sub

Private Sub IREMSOYLU_Click()

IremBilgi.Show

End Sub

Private Sub Sefa_Click()
SefaBilgi.Show
End Sub

Private Sub ToggleButton1_Click()
kemalhocabilgi.Show
End Sub
