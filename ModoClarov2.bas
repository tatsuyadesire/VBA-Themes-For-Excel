Attribute VB_Name = "Módulo2"
Sub ModoClaro()

    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        
        With ws.Cells
            .Interior.Pattern = xlNone
            .Font.ColorIndex = xlAutomatic
            .Borders.LineStyle = xlNone
            .FormatConditions.Delete ' <<< ESSA LINHA resolve
        End With
        
        ws.Tab.ColorIndex = xlColorIndexNone
        
        ws.Activate
        ActiveWindow.DisplayGridlines = True
        ActiveWindow.GridlineColorIndex = xlAutomatic
        
    Next ws

    MsgBox "Excel restaurado ao padrão!", vbInformation

End Sub
