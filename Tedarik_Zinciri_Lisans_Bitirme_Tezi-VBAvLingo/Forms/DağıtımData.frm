VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DağıtımData 
   Caption         =   "Dağıtım Ağı Tasarımı"
   ClientHeight    =   5610
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8055
   OleObjectBlob   =   "DağıtımData.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "DağıtımData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ToggleButton1_Click()
KomponentTuru11.Show
End Sub


Private Sub ToggleButton11_Click()
sınırlar11.Show
End Sub

Private Sub ToggleButton4_Click()
Kapasıteler11.Show
End Sub

Private Sub ToggleButton5_Click()
SabıtMalıyet11.Show
End Sub

Private Sub ToggleButton9_Click()
BırımMalıyet11.Show
End Sub
