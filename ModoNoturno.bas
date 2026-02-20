Attribute VB_Name = "Módulo1"
Sub ModoNoturno()

    Dim ws As Worksheet
    Dim c As Range
    Dim b As Border
    
    For Each ws In ThisWorkbook.Worksheets
        
        ws.Activate
        
        ' Fundo escuro na planilha inteira
        ws.Cells.Interior.Color = RGB(30, 30, 30)
        
        ' Fonte clara
        ws.Cells.Font.Color = RGB(200, 200, 200)
        
        ' Remover gridlines
        ActiveWindow.DisplayGridlines = False
        
        ' Apenas alterar cor das bordas que já existem
        For Each c In ws.UsedRange
            For Each b In c.Borders
                If b.LineStyle <> xlNone Then
                    b.Color = RGB(140, 140, 140)
                End If
            Next b
        Next c
        
        ws.Tab.Color = RGB(60, 60, 60)
        
    Next ws

    MsgBox "Modo Noturno aplicado corretamente!", vbInformation

End Sub
