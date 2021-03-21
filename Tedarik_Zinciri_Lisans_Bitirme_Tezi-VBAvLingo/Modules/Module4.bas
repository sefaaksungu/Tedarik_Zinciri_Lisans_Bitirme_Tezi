Attribute VB_Name = "Module4"
Sub ROTALAMA4()

Sheets("X").Select

Cells(28, 1) = "KONYA"

X = 5
For i = 2 To 34 Step 2
    For j = 2 To 58
           If Cells(X, j) = 1 Then
            Cells(28, i + 1) = Cells(X - (X - 1), j).Value
            a = Cells(X - (X - 1), j).Value
            If a = "KONYA" Then
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
    Range("A28:AG28").Select
    Selection.Copy
    Sheets("Rotalama").Select
    Range("A30").Select
    ActiveSheet.Paste
    
    Range("A30:AF30,A30,C30,E30,G30,I30,K30,M30,O30,Q30,S30,U30,W30,Y30,AA30,AC30,AE30,AG30").Select
    Range("S30").Activate
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
    
    Range("C30:AG30").Select
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
If Cells(28, j).Value = "MANÝSA" Then
Sheets("Rotalama").Select
ActiveSheet.Shapes.Range(Array("Oval 130")).Fill.ForeColor.RGB = vbGreen
Sheets("X").Select
ElseIf Cells(28, j).Value = "EDÝRNE" Then
Sheets("Rotalama").Select
ActiveSheet.Shapes("Oval 8").Fill.ForeColor.RGB = vbGreen
Sheets("X").Select
ElseIf Cells(28, j).Value = "ESKÝÞEHÝR" Then
Sheets("Rotalama").Select
ActiveSheet.Shapes("Oval 131").Fill.ForeColor.RGB = vbGreen
Sheets("X").Select
ElseIf Cells(28, j).Value = "ERZURUM" Then
Sheets("Rotalama").Select
ActiveSheet.Shapes("Oval 141").Fill.ForeColor.RGB = vbGreen
Sheets("X").Select
ElseIf Cells(28, j).Value = "SAMSUN" Then
Sheets("Rotalama").Select
ActiveSheet.Shapes("Oval 133").Fill.ForeColor.RGB = vbGreen
Sheets("X").Select
ElseIf Cells(28, j).Value = "HATAY" Then
Sheets("Rotalama").Select
ActiveSheet.Shapes("Oval 139").Fill.ForeColor.RGB = vbGreen
 Sheets("X").Select
ElseIf Cells(28, j).Value = "SÝVAS" Then
Sheets("Rotalama").Select
ActiveSheet.Shapes("Oval 136").Fill.ForeColor.RGB = vbGreen
 Sheets("X").Select
ElseIf Cells(28, j).Value = "YOZGAT" Then
Sheets("Rotalama").Select
ActiveSheet.Shapes("Oval 135").Fill.ForeColor.RGB = vbGreen
Sheets("X").Select
ElseIf Cells(28, j).Value = "TRABZON" Then
Sheets("Rotalama").Select
ActiveSheet.Shapes("Oval 144").Fill.ForeColor.RGB = vbGreen
 Sheets("X").Select
ElseIf Cells(28, j).Value = "ZONGULDAK" Then
Sheets("Rotalama").Select
ActiveSheet.Shapes("Oval 132").Fill.ForeColor.RGB = vbGreen
 Sheets("X").Select
ElseIf Cells(28, j).Value = "VAN" Then
Sheets("Rotalama").Select
ActiveSheet.Shapes("Oval 143").Fill.ForeColor.RGB = vbGreen
 Sheets("X").Select
ElseIf Cells(28, j).Value = "ÞANLIURFA" Then
Sheets("Rotalama").Select
ActiveSheet.Shapes("Oval 140").Fill.ForeColor.RGB = vbGreen
 Sheets("X").Select
ElseIf Cells(28, j).Value = "KARS" Then
Sheets("Rotalama").Select
ActiveSheet.Shapes("Oval 142").Fill.ForeColor.RGB = vbGreen
 Sheets("X").Select
ElseIf Cells(28, j).Value = "ÇANAKKALE" Then
Sheets("Rotalama").Select
ActiveSheet.Shapes("Oval 128").Fill.ForeColor.RGB = vbGreen
 Sheets("X").Select
ElseIf Cells(28, j).Value = "KAYSERÝ" Then
Sheets("Rotalama").Select
ActiveSheet.Shapes("Oval 137").Fill.ForeColor.RGB = vbGreen
End If
Next j
Sheets("Rotalama").Select
If Cells(30, 3).Value = "" Then
MsgBox "4. Daðýtým Merkezi {KONYA} açýlmamýþtýr ve rotasý oluþmamýþtýr. "
End If
Sheets("Rotalama").Select

For i = 3 To 33
    If Cells(30, i).Value = "KONYA" Then
        Cells(30, i + 2).Select
        k = ActiveCell.Column
        k = k - 1
        For j = k To 33
        k = k + 1
        Cells(30, k).ClearContents
        Next j
    End If
Next i
    Range("J3").Select
    
End Sub

