Sub Macro1()
'
' Macro1 Macro
' eMotion
'
' Atalho do teclado: Ctrl+p
'

' Colocando e formatando o texto na célula
Range("H9").Select
ActiveCell.FormulaR1C1 = "Quantidade de Linhas"
Columns("H:H").ColumnWidth = 13#
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

' Contando a quantidade de objetos
lin = 9
counter = 0
Do Until Cells(lin, 6) = ""
    counter = counter + 1
    lin = lin + 1
Loop
    
Range("H10").Select
ActiveCell.FormulaR1C1 = counter
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