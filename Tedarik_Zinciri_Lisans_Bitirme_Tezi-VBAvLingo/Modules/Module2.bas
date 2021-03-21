Attribute VB_Name = "Module2"
Sub ROTALAMA2()

Sheets("X").Select
Cells(24, 1) = "ANKARA"

X = 3
For i = 2 To 34 Step 2
    For j = 2 To 58
        If Cells(X, j) = 1 Then
            Cells(24, i + 1) = Cells(X - (X - 1), j).Value
            a = Cells(X - (X - 1), j).Value
            If a = "ANKARA" Then
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
    Range("A24:AG24").Select
    Selection.Copy
    Sheets("Rotalama").Select
    Range("A22").Select
    ActiveSheet.Paste
    
    Range("A22:AF22,A22,C22,E22,G22,I22,K22,M22,O22,Q22,S22,U22,W22,Y22,AA22,AC22,AE22,AG22").Select
    Range("S22").Activate
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
    
    Range("C22:AG22").Select
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
If Cells(24, j).Value = "MAN�SA" Then
Sheets("Rotalama").Select
ActiveSheet.Shapes.Range(Array("Oval 130")).Fill.ForeColor.RGB = vbBlue
Sheets("X").Select
ElseIf Cells(24, j).Value = "ED�RNE" Then
Sheets("Rotalama").Select
ActiveSheet.Shapes("Oval 8").Fill.ForeColor.RGB = vbBlue
Sheets("X").Select
ElseIf Cells(24, j).Value = "ESK��EH�R" Then
Sheets("Rotalama").Select
ActiveSheet.Shapes("Oval 131").Fill.ForeColor.RGB = vbBlue
Sheets("X").Select
ElseIf Cells(24, j).Value = "ERZURUM" Then
Sheets("Rotalama").Select
ActiveSheet.Shapes("Oval 141").Fill.ForeColor.RGB = vbBlue
Sheets("X").Select
ElseIf Cells(24, j).Value = "SAMSUN" Then
Sheets("Rotalama").Select
ActiveSheet.Shapes("Oval 133").Fill.ForeColor.RGB = vbBlue
Sheets("X").Select
ElseIf Cells(24, j).Value = "HATAY" Then
Sheets("Rotalama").Select
ActiveSheet.Shapes("Oval 139").Fill.ForeColor.RGB = vbBlue
 Sheets("X").Select
ElseIf Cells(24, j).Value = "S�VAS" Then
Sheets("Rotalama").Select
ActiveSheet.Shapes("Oval 136").Fill.ForeColor.RGB = vbBlue
 Sheets("X").Select
ElseIf Cells(24, j).Value = "YOZGAT" Then
Sheets("Rotalama").Select
ActiveSheet.Shapes("Oval 135").Fill.ForeColor.RGB = vbBlue
Sheets("X").Select
ElseIf Cells(24, j).Value = "TRABZON" Then
Sheets("Rotalama").Select
ActiveSheet.Shapes("Oval 144").Fill.ForeColor.RGB = vbBlue
 Sheets("X").Select
ElseIf Cells(24, j).Value = "ZONGULDAK" Then
Sheets("Rotalama").Select
ActiveSheet.Shapes("Oval 132").Fill.ForeColor.RGB = vbBlue
 Sheets("X").Select
ElseIf Cells(24, j).Value = "VAN" Then
Sheets("Rotalama").Select
ActiveSheet.Shapes("Oval 143").Fill.ForeColor.RGB = vbBlue
 Sheets("X").Select
ElseIf Cells(24, j).Value = "�ANLIURFA" Then
Sheets("Rotalama").Select
ActiveSheet.Shapes("Oval 140").Fill.ForeColor.RGB = vbBlue
 Sheets("X").Select
ElseIf Cells(24, j).Value = "KARS" Then
Sheets("Rotalama").Select
ActiveSheet.Shapes("Oval 142").Fill.ForeColor.RGB = vbBlue
 Sheets("X").Select
ElseIf Cells(24, j).Value = "�ANAKKALE" Then
Sheets("Rotalama").Select
ActiveSheet.Shapes("Oval 128").Fill.ForeColor.RGB = vbBlue
 Sheets("X").Select
ElseIf Cells(24, j).Value = "KAYSER�" Then
Sheets("Rotalama").Select
ActiveSheet.Shapes("Oval 137").Fill.ForeColor.RGB = vbBlue
End If
Next j
Sheets("Rotalama").Select
If Cells(22, 3).Value = "" Then
MsgBox "2. Da��t�m Merkezi {ANKARA} a��lmam��t�r ve rotas� olu�mam��t�r. "
End If
Sheets("Rotalama").Select

For i = 3 To 33
    If Cells(22, i).Value = "ANKARA" Then
        Cells(22, i + 2).Select
        k = ActiveCell.Column
        k = k - 1
        For j = k To 33
        k = k + 1
        Cells(22, k).ClearContents
        Next j
    End If
Next i
    Range("J3").Select
End Sub

