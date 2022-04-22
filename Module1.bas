Attribute VB_Name = "Module1"
Sub AddHeaders()
Attribute AddHeaders.VB_Description = "Automate adding headers to the worksheet"
Attribute AddHeaders.VB_ProcData.VB_Invoke_Func = "j\n14"
' AddHeaders Macro
' Automate adding headers to the worksheet
'
' Keyboard Shortcut: Ctrl+j
'
'Cuando grabamos la macro:
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "Region"
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "Expense"
    Range("C3").Select
    ActiveCell.FormulaR1C1 = "Jan"
    Range("D3").Select
    ActiveCell.FormulaR1C1 = "Feb"
    Range("E3").Select
    ActiveCell.FormulaR1C1 = "Mar"
    Range("F3").Select
    ActiveCell.FormulaR1C1 = "Total"
    Range("A3:F3").Select
    Application.Left = 49
    Application.Top = 211
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Application.Left = 302
    Application.Top = 133
    Selection.Font.Bold = True
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Application.Left = 328.5
    Application.Top = 53
End Sub

