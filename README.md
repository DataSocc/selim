Option Explicit

Public Sub ImportarSelicSemLibJson()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("selic2")
    
    '-----------------------------------------------
    ' 1) Limpar da linha 2 até o final (colunas A:C)
    '-----------------------------------------------
    ws.Range("A2:C100000").ClearContents
    
    '-----------------------------------------------
    ' 2) Calcular hoje e data inicial (5 anos atrás)
    '-----------------------------------------------
    Dim dtHoje As Date
    dtHoje = Date
    
    Dim dtInicial As Date
    dtInicial = DateAdd("yyyy", -5, dtHoje)
    
    '-----------------------------------------------
    ' 3) Montar o body do POST (JSON simples)
    '-----------------------------------------------
    Dim postBody As String
    postBody = "{""dataInicial"":""" & Format(dtInicial, "dd/MM/yyyy") & """," & _
               """dataFinal"":""" & Format(dtHoje, "dd/MM/yyyy") & """}"
    
    '-----------------------------------------------
    ' 4) Fazer o POST com MSXML2.XMLHTTP
    '-----------------------------------------------
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    With http
        .Open "POST", _
              "https://www3.bcb.gov.br/novoselic/rest/taxaSelicApurada/pub/search?parametrosOrdenacao=%5B%5D&page=1&pageSize=999999999", _
              False
        .setRequestHeader "Accept-Language", "pt-BR"   ' Exemplo: se precisar
        .setRequestHeader "Content-Type", "application/json"
        .send (postBody)
    End With
    
    Dim jsonResponse As String
    jsonResponse = http.responseText
    
    '----------------------------------------------------------------------------
    ' 5) Parsear MANUALMENTE o JSON para "dataCotacao" e "fatorDiario"
    '
    '    - Procuramos por trechos como:
    '        "dataCotacao": "30/01/2025"
    '        "fatorDiario": 1.00049037
    '----------------------------------------------------------------------------
    
    Dim pos As Long
    pos = 1
    
    Dim rowOut As Long
    rowOut = 2
    
    Dim countFound As Long
    countFound = 0
    
    Dim posDataCotacao As Long
    Dim posColon As Long
    Dim posDataValueStart As Long, posDataValueEnd As Long
    
    Dim txtDataCotacao As String
    Dim txtFator As String
    
    Dim posFator As Long
    Dim posFatorColon As Long
    Dim posFatorValueStart As Long, posFatorValueEnd As Long
    
    Do
        ' (a) Achar a próxima ocorrência de "dataCotacao":
        posDataCotacao = InStr(pos, jsonResponse, """dataCotacao"":", vbTextCompare)
        If posDataCotacao = 0 Then Exit Do  ' não encontrou, paramos
        
        ' Localiza o ':' após "dataCotacao"
        posColon = InStr(posDataCotacao, jsonResponse, ":", vbTextCompare)
        If posColon = 0 Then Exit Do
        
        ' Localiza as aspas que abrem a data
        posDataValueStart = InStr(posColon, jsonResponse, """", vbTextCompare)
        If posDataValueStart = 0 Then Exit Do
        
        ' Localiza as aspas que fecham a data
        posDataValueEnd = InStr(posDataValueStart + 1, jsonResponse, """", vbTextCompare)
        If posDataValueEnd = 0 Then Exit Do
        
        txtDataCotacao = Mid(jsonResponse, posDataValueStart + 1, _
                             posDataValueEnd - (posDataValueStart + 1))
        
        ' (b) Procurar "fatorDiario" a partir de posDataValueEnd
        posFator = InStr(posDataValueEnd, jsonResponse, """fatorDiario"":", vbTextCompare)
        If posFator = 0 Then Exit Do
        
        posFatorColon = InStr(posFator, jsonResponse, ":", vbTextCompare)
        If posFatorColon = 0 Then Exit Do
        
        ' Agora procuramos o número do fator (sem aspas), 
        ' então chamamos a função auxiliar para extrair o valor numérico
        txtFator = ExtrairValorNumerico(jsonResponse, posFatorColon + 1)
        
        ' (c) Preencher na planilha
        ws.Cells(rowOut, 1).Value = txtDataCotacao   ' DataCotacao (string)
        ws.Cells(rowOut, 1).NumberFormat = "dd/mm/yyyy"
        
        ws.Cells(rowOut, 2).Value = Val(txtFator)    ' fatorDiario (double)
        ws.Cells(rowOut, 2).NumberFormat = "#,##0.00000000"
        
        ws.Cells(rowOut, 3).Value = (rowOut - 1)     ' ID sequencial (1,2,3,...)
        
        rowOut = rowOut + 1
        countFound = countFound + 1
        
        ' Avança 'pos' para procurar o próximo registro
        pos = posFator + 1
        
    Loop
    
    MsgBox "Processo concluído! " & countFound & " linhas inseridas.", vbInformation

End Sub

' ----------------------------------------------------------------------------
' Função auxiliar que lê a partir de 'startPos' e extrai somente a parte
' numérica (podendo ter ponto, vírgula ou sinal -), parando no primeiro char
' inválido (por exemplo, ',', '}') 
' ----------------------------------------------------------------------------
Private Function ExtrairValorNumerico(ByVal source As String, ByVal startPos As Long) As String
    Dim temp As String
    temp = Mid(source, startPos, 50)  ' lê até 50 chars (ajuste se precisar)
    
    temp = LTrim(temp)               ' remove espaços à esquerda
    
    Dim i As Long
    Dim ch As String
    Dim ret As String
    
    For i = 1 To Len(temp)
        ch = Mid(temp, i, 1)
        
        ' Aceitamos dígitos 0-9, ponto (.) ou vírgula (,) e '-' se for negativo
        If (ch Like "[0-9]") Or (ch = ".") Or (ch = ",") Or (ch = "-") Then
            ret = ret & ch
        Else
            Exit For
        End If
    Next i
    
    ' Se houver vírgulas decimais, trocamos por ponto
    ret = Replace(ret, ",", ".")
    
    ExtrairValorNumerico = ret
End Function
