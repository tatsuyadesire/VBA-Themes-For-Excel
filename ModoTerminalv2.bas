Attribute VB_Name = "Módulo3"
Sub ModoTerminal()

    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        
        With ws.Cells
            
            .FormatConditions.Delete
            .ClearFormats
            
            ' Fundo preto
            .Interior.Color = RGB(0, 0, 0)
            
            ' Fonte verde Matrix
            .Font.Color = RGB(0, 255, 70)
            
        End With
        
        ws.Tab.Color = RGB(0, 0, 0)
        
        ws.Activate
        ActiveWindow.DisplayGridlines = False
        
    Next ws

    MsgBox "Modo Terminal ativado", vbInformation

End Sub
