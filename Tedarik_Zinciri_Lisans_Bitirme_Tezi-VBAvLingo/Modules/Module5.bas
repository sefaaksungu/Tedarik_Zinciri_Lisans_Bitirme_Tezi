Attribute VB_Name = "Module5"

Sub RotaSil()
Attribute RotaSil.VB_ProcData.VB_Invoke_Func = " \n14"

    ActiveWindow.SmallScroll Down:=-11
    Range("A15:AG32").Select
    Selection.ClearContents
    Range("A15").Select
    Sheets("X").Select
    Range("A22:AG28").Select
    Selection.ClearContents
    Range("A30").Select
    Sheets("Rotalama").Select
    Range("J9").Select
    
    ActiveSheet.Shapes.Range(Array("Oval 8")).Select
    ActiveSheet.Shapes.Range(Array("Oval 8", "Oval 128")).Select
    ActiveSheet.Shapes.Range(Array("Oval 8", "Oval 128", "Oval 130")).Select
    ActiveSheet.Shapes.Range(Array("Oval 8", "Oval 128", "Oval 130", "Oval 132" _
        )).Select
    ActiveSheet.Shapes.Range(Array("Oval 8", "Oval 128", "Oval 130", "Oval 132", _
        "Oval 131")).Select
    ActiveSheet.Shapes.Range(Array("Oval 8", "Oval 128", "Oval 130", "Oval 132", _
        "Oval 131", "Oval 135")).Select
    ActiveSheet.Shapes.Range(Array("Oval 8", "Oval 128", "Oval 130", "Oval 132", _
        "Oval 131", "Oval 135", "Oval 133")).Select
    ActiveSheet.Shapes.Range(Array("Oval 8", "Oval 128", "Oval 130", "Oval 132", _
        "Oval 131", "Oval 135", "Oval 133", "Oval 137")).Select
    ActiveSheet.Shapes.Range(Array("Oval 8", "Oval 128", "Oval 130", "Oval 132", _
        "Oval 131", "Oval 135", "Oval 133", "Oval 137", "Oval 139")).Select
    ActiveSheet.Shapes.Range(Array("Oval 8", "Oval 128", "Oval 130", "Oval 132", _
        "Oval 131", "Oval 135", "Oval 133", "Oval 137", "Oval 139", "Oval 136")). _
        Select
    ActiveSheet.Shapes.Range(Array("Oval 8", "Oval 128", "Oval 130", "Oval 132", _
        "Oval 131", "Oval 135", "Oval 133", "Oval 137", "Oval 139", "Oval 136", _
        "Oval 140")).Select
    ActiveSheet.Shapes.Range(Array("Oval 8", "Oval 128", "Oval 130", "Oval 132", _
        "Oval 131", "Oval 135", "Oval 133", "Oval 137", "Oval 139", "Oval 136", _
        "Oval 140", "Oval 144")).Select
    ActiveSheet.Shapes.Range(Array("Oval 8", "Oval 128", "Oval 130", "Oval 132", _
        "Oval 131", "Oval 135", "Oval 133", "Oval 137", "Oval 139", "Oval 136", _
        "Oval 140", "Oval 144", "Oval 141")).Select
    ActiveSheet.Shapes.Range(Array("Oval 8", "Oval 128", "Oval 130", "Oval 132", _
        "Oval 131", "Oval 135", "Oval 133", "Oval 137", "Oval 139", "Oval 136", _
        "Oval 140", "Oval 144", "Oval 141", "Oval 142")).Select
    ActiveSheet.Shapes.Range(Array("Oval 8", "Oval 128", "Oval 130", "Oval 132", _
        "Oval 131", "Oval 135", "Oval 133", "Oval 137", "Oval 139", "Oval 136", _
        "Oval 140", "Oval 144", "Oval 141", "Oval 142", "Oval 143")).Select
    With Selection.ShapeRange.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 255, 255)
        .Transparency = 0
        .Solid
    End With
    Range("J9").Select
End Sub
