Attribute VB_Name = "Module1"
Sub ROTALAMA1()
Sheets("X").Select
Cells(22, 1) = "ÝZMÝR"

X = 2
For i = 2 To 34 Step 2
    For j = 2 To 58
        If Cells(X, j) = 1 Then
            Cells(22, i + 1) = Cells(X - (X - 1), j).Value
            a = Cells(X - (X - 1), j).Value
            If a = "ÝZMÝR" Then
            Exit For
            End If
            For k = 2 To 20
                   If Cells(k, 1) = a Then
                      Cells(k, 1).Select
                      m = ActiveCell.Row
                      X = m
                      Exit For
                   End If
            Next k
            Exit For
        End If
    Next j
Next i
    Sheets("X").Select
    Range("A22:AG22").Select
    Selection.Copy
    Sheets("Rotalama").Select
    Range("A18").Select
    ActiveSheet.Paste
    
    Range("A18:AF18,A18,C18,E18,G18,I18,K18,M18,O18,Q18,S18,U18,W18,Y18,AA18,AC18,AE18,AG18").Select
    Range("S18").Activate
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
    Range("C18:AG18").Select
    Selection.FormatConditions.AddUniqueValues
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).DupeUnique = xlDuplicate
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 4.99893185216834E-02
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Cells(3, 10).Select
    
    Sheets("Rotalama").Select
    Range("J3").Select
    
Sheets("X").Select
For j = 3 To 33
Sheets("X").Select
If Cells(22, j).Value = "MANÝSA" Then
Sheets("Rotalama").Select
ActiveSheet.Shapes.Range(Array("Oval 130")).Fill.ForeColor.RGB = vbRed
Sheets("X").Select
ElseIf Cells(22, j).Value = "EDÝRNE" Then
Sheets("Rotalama").Select
ActiveSheet.Shapes("Oval 8").Fill.ForeColor.RGB = vbRed
Sheets("X").Select
ElseIf Cells(22, j).Value = "ESKÝÞEHÝR" Then
Sheets("Rotalama").Select
ActiveSheet.Shapes("Oval 131").Fill.ForeColor.RGB = vbRed
Sheets("X").Select
ElseIf Cells(22, j).Value = "ERZURUM" Then
Sheets("Rotalama").Select
ActiveSheet.Shapes("Oval 141").Fill.ForeColor.RGB = vbRed
Sheets("X").Select
ElseIf Cells(22, j).Value = "SAMSUN" Then
Sheets("Rotalama").Select
ActiveSheet.Shapes("Oval 133").Fill.ForeColor.RGB = vbRed
Sheets("X").Select
ElseIf Cells(22, j).Value = "HATAY" Then
Sheets("Rotalama").Select
ActiveSheet.Shapes("Oval 139").Fill.ForeColor.RGB = vbRed
 Sheets("X").Select
ElseIf Cells(22, j).Value = "SÝVAS" Then
Sheets("Rotalama").Select
ActiveSheet.Shapes("Oval 136").Fill.ForeColor.RGB = vbRed
 Sheets("X").Select
ElseIf Cells(22, j).Value = "YOZGAT" Then
Sheets("Rotalama").Select
ActiveSheet.Shapes("Oval 135").Fill.ForeColor.RGB = vbRed
Sheets("X").Select
ElseIf Cells(22, j).Value = "TRABZON" Then
Sheets("Rotalama").Select
ActiveSheet.Shapes("Oval 144").Fill.ForeColor.RGB = vbRed
 Sheets("X").Select
ElseIf Cells(22, j).Value = "ZONGULDAK" Then
Sheets("Rotalama").Select
ActiveSheet.Shapes("Oval 132").Fill.ForeColor.RGB = vbRed
 Sheets("X").Select
ElseIf Cells(22, j).Value = "VAN" Then
Sheets("Rotalama").Select
ActiveSheet.Shapes("Oval 143").Fill.ForeColor.RGB = vbRed
 Sheets("X").Select
ElseIf Cells(22, j).Value = "ÞANLIURFA" Then
Sheets("Rotalama").Select
ActiveSheet.Shapes("Oval 140").Fill.ForeColor.RGB = vbRed
 Sheets("X").Select
ElseIf Cells(22, j).Value = "KARS" Then
Sheets("Rotalama").Select
ActiveSheet.Shapes("Oval 142").Fill.ForeColor.RGB = vbRed
 Sheets("X").Select
ElseIf Cells(22, j).Value = "ÇANAKKALE" Then
Sheets("Rotalama").Select
ActiveSheet.Shapes("Oval 128").Fill.ForeColor.RGB = vbRed
 Sheets("X").Select
ElseIf Cells(22, j).Value = "KAYSERÝ" Then
Sheets("Rotalama").Select
ActiveSheet.Shapes("Oval 137").Fill.ForeColor.RGB = vbRed
End If
Next j
Sheets("Rotalama").Select
If Cells(18, 3).Value = "" Then
MsgBox "1. Daðýtým Merkezi {ÝZMÝR} açýlmamýþtýr ve rotasý oluþmamýþtýr. "
End If
Sheets("Rotalama").Select

For i = 3 To 33
    If Cells(18, i).Value = "ÝZMÝR" Then
        Cells(18, i + 2).Select
        k = ActiveCell.Column
        k = k - 1
        For j = k To 33
        k = k + 1
        Cells(18, k).ClearContents
        Next j
    End If
Next i
    Range("J3").Select
End Sub


