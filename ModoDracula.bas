Attribute VB_Name = "Módulo4"
Sub DraculaTheme()

    Dim ws As Worksheet
    Dim fc As FormatCondition
    
    Application.ScreenUpdating = False
    
    For Each ws In ThisWorkbook.Worksheets
        
        ws.Cells.Interior.Color = RGB(40, 42, 54)
        ws.Cells.Font.Name = "Consolas"
        ws.Cells.Font.Size = 11
        ws.Cells.Font.Color = RGB(248, 248, 242)
        
        ActiveWindow.DisplayGridlines = True
        ActiveWindow.GridlineColor = RGB(68, 71, 90)
        
        ws.Cells.FormatConditions.Delete
        
        ' ERROS (ROSA)
        Set fc = ws.Cells.FormatConditions.Add(Type:=xlExpression, _
            Formula1:="=ÉERROS(A1)")
        fc.Font.Color = RGB(255, 121, 198)
        fc.StopIfTrue = True
        
        ' FÓRMULAS (ROXO)
        Set fc = ws.Cells.FormatConditions.Add(Type:=xlExpression, _
            Formula1:="=ÉFÓRMULA(A1)")
        fc.Font.Color = RGB(189, 147, 249)
        fc.StopIfTrue = True
        
        ' NÚMEROS (VERDE)
        Set fc = ws.Cells.FormatConditions.Add(Type:=xlExpression, _
            Formula1:="=E(ÉNÚM(A1);NÃO(ÉFÓRMULA(A1)))")
        fc.Font.Color = RGB(80, 250, 123)
        fc.StopIfTrue = True
        
        ' TEXTO (AMARELO)
        Set fc = ws.Cells.FormatConditions.Add(Type:=xlExpression, _
            Formula1:="=ÉTEXTO(A1)")
        fc.Font.Color = RGB(241, 250, 140)
        
        ws.Tab.Color = RGB(189, 147, 249)
        
    Next ws
    
    Application.ScreenUpdating = True
    
    MsgBox "Drácula aplicado corretamente", vbInformation

End Sub

