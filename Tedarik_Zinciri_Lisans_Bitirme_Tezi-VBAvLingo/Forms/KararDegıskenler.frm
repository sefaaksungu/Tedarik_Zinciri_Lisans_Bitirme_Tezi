VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} KararDegıskenler 
   Caption         =   "Karar Değişkenleri Grubu"
   ClientHeight    =   11025
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20460
   OleObjectBlob   =   "KararDegıskenler.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "KararDegıskenler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
Sheets("A_1").Select
KararDegıskenler.Hide
ANASAYFA.Hide

End Sub


Private Sub CommandButton2_Click()
Sheets("FS").Select
KararDegıskenler.Hide
ANASAYFA.Hide
End Sub

Private Sub CommandButton3_Click()
Sheets("Rotalama").Select
KararDegıskenler.Hide
ANASAYFA.Hide

End Sub

Private Sub CommandButton4_Click()
Sheets("A_3").Select
KararDegıskenler.Hide
ANASAYFA.Hide

End Sub

Private Sub CommandButton5_Click()
Sheets("A_2").Select
KararDegıskenler.Hide
ANASAYFA.Hide

End Sub

Private Sub CommandButton10_Click()
Sheets("Z").Select
KararDegıskenler.Hide
ANASAYFA.Hide

End Sub

Private Sub CommandButton11_Click()
Sheets("U ve F").Select
KararDegıskenler.Hide
ANASAYFA.Hide

End Sub

Private Sub CommandButton9_Click()
Sheets("P").Select
KararDegıskenler.Hide
ANASAYFA.Hide

End Sub
