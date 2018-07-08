Sub Macro4()
'
' Macro4 Macro
' eMotion 4
'
' Atalho do teclado: Ctrl+u
'

Range("K9").Select
ActiveCell.FormulaR1C1 = "Maior Valor"
Columns("K:K").ColumnWidth = 13#
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

Max = Cells(10, 6)
lin = 11
Do Until Cells(lin, 6) = ""
    
    If Max < Cells(lin, 6) Then
        Max = Cells(lin, 6)
    End If
    Cells(lin, 7) = Max
    lin = lin + 1
Loop

Range("K10").Select
ActiveCell.FormulaR1C1 = Max
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