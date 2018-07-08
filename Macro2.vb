Sub Macro2()
'
' Macro2 Macro
' eMotion 2
'
' Atalho do teclado: Ctrl+ç
'

Range("I9").Select
ActiveCell.FormulaR1C1 = "Célula B9 é Lucas?"
Columns("I:I").ColumnWidth = 13#
Rows("9:9").RowHeight = 30#
Rows("10:10").RowHeight = 15

With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
End With

'A célula que deve ser avaliada é a célula B10
Range("I10").Select
If Range("B10").Text = "Lucas" Then
    ActiveCell.FormulaR1C1 = "SIM"
Else
    ActiveCell.FormulaR1C1 = "NÃO"
End If

With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
End With

End Sub