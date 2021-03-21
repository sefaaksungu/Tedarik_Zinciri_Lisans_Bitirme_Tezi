Attribute VB_Name = "Module7"
Sub PDFOLUSTUR()
    
    sor2 = MsgBox("Raporu PDF Dosyasý Olarak Kaydetmek Ýstiyor Musunuz ? ", vbYesNo)
    If sor2 = vbNo Then Exit Sub
    
    yol = Sheets("Rapor").Cells(42, 4)
    
    
    isim = Cells(12, 15)
    
    Dim Fs As Object
    Set Fs = CreateObject("Scripting.FileSystemObject")
    If Fs.FileExists(yol & "/" & isim & ".pdf") Then
        sor = MsgBox("Dosya Var Yinede Devam Etmek Ýstiyor Musunuz ? ", vbYesNo)
        If sor = vbNo Then Exit Sub
    Else
    End If
    
    Sheets("Rapor").Select
    ActiveSheet.Range("A1:L41").ExportAsFixedFormat Type:=xlTypePDF, Filename:=yol & "/" & isim & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True
    
End Sub

