Attribute VB_Name = "Module60"
Sub talephesabý()
Sheets("DATA {1}").Select
toplam = 0
For i = 7 To 21
    toplam = toplam + Cells(23, i)
Next i
    If toplam > 7500 Then
    Sheets("KARAR DESTEK").Select
    MsgBox "Ýllere ait girilmiþ olunan talepler, kapasite aþýmýna sebep olmaktadýr. Lütfen talepleri yeniden düzenleyiniz. "
    Sheets("KARAR DESTEK").ManisaTalep.Value = 0
    Worksheets("DATA {1}").Cells(23, 7) = 0
    Sheets("KARAR DESTEK").EdirneTalep.Value = 0
    Worksheets("DATA {1}").Cells(23, 8) = 0
    Sheets("KARAR DESTEK").EskiþehirTalep.Value = 0
    Worksheets("DATA {1}").Cells(23, 9) = 0
    Sheets("KARAR DESTEK").ErzurumTalep.Value = 0
    Worksheets("DATA {1}").Cells(23, 10) = 0
    Sheets("KARAR DESTEK").SamsunTalep.Value = 0
    Worksheets("DATA {1}").Cells(23, 11) = 0
    Sheets("KARAR DESTEK").HatayTalep.Value = 0
    Worksheets("DATA {1}").Cells(23, 12) = 0
    Sheets("KARAR DESTEK").SivasTalep.Value = 0
    Worksheets("DATA {1}").Cells(23, 13) = 0
    Sheets("KARAR DESTEK").YozgatTalep.Value = 0
    Worksheets("DATA {1}").Cells(23, 14) = 0
    Sheets("KARAR DESTEK").TrabzonTalep.Value = 0
    Worksheets("DATA {1}").Cells(23, 15) = 0
    Sheets("KARAR DESTEK").ZonguldakTalep.Value = 0
    Worksheets("DATA {1}").Cells(23, 16) = 0
    Sheets("KARAR DESTEK").VanTalep.Value = 0
    Worksheets("DATA {1}").Cells(23, 17) = 0
    Sheets("KARAR DESTEK").ÞanlýurfaTalep.Value = 0
    Worksheets("DATA {1}").Cells(23, 18) = 0
    Sheets("KARAR DESTEK").KarsTalep.Value = 0
    Worksheets("DATA {1}").Cells(23, 19) = 0
    Sheets("KARAR DESTEK").ÇanakkaleTalep.Value = 0
    Worksheets("DATA {1}").Cells(23, 20) = 0
    Sheets("KARAR DESTEK").KayseriTalep.Value = 0
    Worksheets("DATA {1}").Cells(23, 21) = 0
    Else
    Sheets("KARAR DESTEK").Select
    MsgBox "Talepleriniz onaylanýp girdi olarak saðlanmýþtýr. Artýk 1. Aþamayý çözdürebilirsiniz."
    End If
End Sub
