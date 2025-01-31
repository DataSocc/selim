Sub ReplicarTaxaProximos5Anos()
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim novaLinha As Long
    Dim dataInicial As Date
    Dim taxa As Double
    Dim i As Long
    
    ' Definir a planilha "Selic"
    Set ws = ThisWorkbook.Sheets("Selic")
    
    ' Encontrar a última linha preenchida na coluna A (Datas)
    ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Pegar a última data e a última taxa registradas
    dataInicial = ws.Cells(ultimaLinha, 1).Value
    taxa = ws.Cells(ultimaLinha, 2).Value
    
    ' Definir a nova linha de inserção
    novaLinha = ultimaLinha + 1
    
    ' Loop para adicionar os próximos 5 anos (1 dia por vez)
    For i = 1 To 5 * 365 ' 5 anos * 365 dias
        ws.Cells(novaLinha, 1).Value = dataInicial + i ' Adiciona 1 dia a cada iteração
        ws.Cells(novaLinha, 2).Value = taxa ' Copia a taxa
        novaLinha = novaLinha + 1 ' Próxima linha
    Next i
    
    ' Formatar a coluna de datas para "dd/mm/yyyy"
    ws.Columns(1).NumberFormat = "dd/mm/yyyy"
    
    MsgBox "Taxa replicada para os próximos 5 anos!", vbInformation, "Concluído"
End Sub
