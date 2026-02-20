Attribute VB_Name = "Módulo3"
Sub ModoTerminal()

    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        
        ws.Activate
        
        ' Fundo preto
        ws.Cells.Interior.Color = RGB(0, 0, 0)
        
        ' Fonte verde terminal
        ws.Cells.Font.Color = RGB(0, 255, 0)
        ws.Cells.Font.Name = "Consolas"
        ws.Cells.Font.Size = 11
        
        ' Manter gridlines desligadas
        ActiveWindow.DisplayGridlines = False
        
        ' Alterar cor das bordas existentes para verde escuro
        Dim c As Range
        Dim b As Border
        
        For Each c In ws.UsedRange
            For Each b In c.Borders
                If b.LineStyle <> xlNone Then
                    b.Color = RGB(0, 150, 0)
                End If
            Next b
        Next c
        
        ws.Tab.Color = RGB(0, 0, 0)
        
    Next ws
    
    MsgBox "Modo Terminal Ativado", vbInformation

End Sub
