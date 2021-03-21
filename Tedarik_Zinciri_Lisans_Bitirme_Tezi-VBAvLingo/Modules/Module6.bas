Attribute VB_Name = "Module6"

Sub RaporOluþtur()
Attribute RaporOluþtur.VB_ProcData.VB_Invoke_Func = " \n14"

    Sheets("Rotalama").Select
    ActiveWindow.ScrollColumn = 1
    Range("C18:AG18").Select
    Selection.Copy
    Sheets("Rapor").Select
    Range("E6").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
    ActiveWindow.SmallScroll Down:=11
    Range("B31").Select
    Sheets("Rotalama").Select
    ActiveWindow.ScrollColumn = 1
    Range("C22:AG22").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Rapor").Select
    ActiveWindow.SmallScroll Down:=-11
    Range("G6").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
    Sheets("Rotalama").Select
    ActiveWindow.ScrollColumn = 1
    Range("C26:AG26").Select
    Application.CutCopyMode = False
    Selection.Copy
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Rapor").Select
    Range("I6").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
    ActiveWindow.SmallScroll Down:=0
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
    Range("C25").Select
    Sheets("Rotalama").Select
    ActiveWindow.ScrollColumn = 1
    Range("C30:AG30").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Rapor").Select
    Range("K6").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
    Range("B24").Select
    ActiveWindow.SmallScroll Down:=11
    Range("F35,H35,J35,L35").Select
    Range("L35").Activate
    Application.CutCopyMode = False
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    ActiveWindow.SmallScroll Down:=11
    Range("A1:L43").Select
    Range("L43").Activate
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
    ActiveWindow.SmallScroll Down:=-22
    Range("B25").Select
    ActiveWindow.SmallScroll Down:=-11
    Sheets("DATA {1}").Select
    Range("G23:U23").Select
    Selection.Copy
    Sheets("Rapor").Select
    Range("C8").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
    Range("C8:C22").Select
    Application.CutCopyMode = False
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
    Range("O25").Select
   
    
    Range("B28").Select
    ActiveCell.FormulaR1C1 = _
        "=SUMPRODUCT(TCOST1,A_1)+SUMPRODUCT(TCOST2,A_2)+SUMPRODUCT(TCOST3,A_3)+SUMPRODUCT(FCP,U)+SUMPRODUCT(FCD,F)"
    Range("B32").Select
    ActiveCell.FormulaR1C1 = "=SUMPRODUCT(DstanceCT,X)"
    Range("B36").Select
    ActiveCell.FormulaR1C1 = "=SUMPRODUCT(FCFS*FS)"
    Range("B39").Select
    ActiveCell.FormulaR1C1 = "=R[-11]C+R[-7]C+R[-3]C"
    
    
    
    Union(Range( _
        "K10,K12,K14,K16,K18,K20,K22,K24,K26,I26,G26,E26,E28,G28,I28,K28,E6,G6,I6,K6,K8,I8,G8,E8,E10,G10,E12,G12,E14,G14,E16,G16" _
        ), Range("E18,G18,E20,G20,E22,G22,E24,G24,I24,I22,I20,I18,I16,I14,I12,I10")). _
        Select
    Range("K28").Activate
    ActiveWindow.ScrollRow = 3
    ActiveWindow.ScrollRow = 4
    ActiveWindow.ScrollRow = 5
    ActiveWindow.ScrollRow = 7
    ActiveWindow.ScrollRow = 8
    ActiveWindow.ScrollRow = 9
    ActiveWindow.ScrollRow = 10
    ActiveWindow.ScrollRow = 11
    ActiveWindow.ScrollRow = 12
    ActiveWindow.ScrollRow = 13
    ActiveWindow.ScrollRow = 14
    Union(Range( _
        "K10,K12,K14,K16,K18,K20,K22,K24,K26,I26,G26,E26,E28,G28,I28,K28,K30,I30,I32,K32,I34,K34,I36,K36,G36,G34,G32,G30,E30,E32,E34,E36" _
        ), Range( _
        "E6,G6,I6,K6,K8,I8,G8,E8,E10,G10,E12,G12,E14,G14,E16,G16,E18,G18,E20,G20,E22,G22,E24,G24,I24,I22,I20,I18,I16,I14,I12,I10" _
        )).Select
    Range("E36").Activate
    ActiveWindow.ScrollRow = 13
    ActiveWindow.ScrollRow = 12
    ActiveWindow.ScrollRow = 11
    ActiveWindow.ScrollRow = 10
    ActiveWindow.ScrollRow = 9
    ActiveWindow.ScrollRow = 8
    ActiveWindow.ScrollRow = 7
    ActiveWindow.ScrollRow = 6
    ActiveWindow.ScrollRow = 5
    ActiveWindow.ScrollRow = 4
    ActiveWindow.ScrollRow = 3
    ActiveWindow.ScrollRow = 2
    ActiveWindow.ScrollRow = 1
    Union(Range( _
        "K10,K12,K14,K16,K18,K20,K22,K24,K26,I26,G26,E26,E28,G28,I28,K28,K30,I30,I32,K32,I34,K34,I36,K36,G36,G34,G32,G30,E30,E32,E34,E36" _
        ), Range( _
        "C8,C8:C22,E6,G6,I6,K6,K8,I8,G8,E8,E10,G10,E12,G12,E14,G14,E16,G16,E18,G18,E20,G20,E22,G22,E24,G24,I24,I22,I20,I18,I16,I14" _
        ), Range("I12,I10")).Select
    Range("C8").Activate
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("B8:C22").Select
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
    Range("C8:C22").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    ActiveWindow.SmallScroll Down:=11
    Range("B40:I40").Select
    ActiveCell.FormulaR1C1 = _
        "Ýletiþim: sefa.aksunguu@gmail.com {05054222415} ve soyluu.iremm@gmail.com "
   
    Range("R4").Select

End Sub
