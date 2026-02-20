Attribute VB_Name = "Módulo2"
Sub ModoClaro()

    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        
        ws.Cells.Interior.Pattern = xlNone
        ws.Cells.Font.ColorIndex = xlAutomatic
        ws.Cells.Borders.LineStyle = xlNone
        
        ws.Tab.ColorIndex = xlColorIndexNone
        
        ws.Activate
        ActiveWindow.DisplayGridlines = True
        ActiveWindow.GridlineColorIndex = xlAutomatic
        
    Next ws

    MsgBox "Excel restaurado ao padrão!", vbInformation

End Sub
