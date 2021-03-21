VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ANASAYFA 
   Caption         =   "ANASAYFA"
   ClientHeight    =   11025
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20460
   OleObjectBlob   =   "ANASAYFA.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "ANASAYFA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
karardestekerisim.Show
End Sub

Private Sub DataveNotasyon_Click()
DataveNotasyonlar.Show
End Sub

Private Sub ÝletiþimBilgileri_Click()
ÝleþimBilgisi.Show
End Sub

Private Sub InsanKaynaklarý_Click()
InsanKaynak.Show
End Sub

Private Sub KararDegýskenlerý_Click()
KararDegýskenler.Show
End Sub

Private Sub ToggleButton12_Click()
amacfonk.Show
End Sub
