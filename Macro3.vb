Sub Macro3()
'
' Macro3 Macro
' eMotion 3
'
' Atalho do teclado: Ctrl+i
'
    Range("C10").Select
    Selection.Copy
    Range("J10").Select
    ActiveSheet.Paste
    
    Columns("J:J").ColumnWidth = 13#
    Rows("10:10").RowHeight = 15
    
    
End Sub