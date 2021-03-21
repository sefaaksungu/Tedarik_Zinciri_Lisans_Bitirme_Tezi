VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AracData 
   Caption         =   "Araç Rotalama"
   ClientHeight    =   5610
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8055
   OleObjectBlob   =   "AracData.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "AracData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ToggleButton1_Click()
MusterýlerArasýMesafe21.Show
End Sub

Private Sub ToggleButton2_Click()
AracKullanmaMalýyetý21.Show
End Sub
