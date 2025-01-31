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
    '    - Ignoramos outras chaves.
    '    - Cada vez que achar "dataCotacao", pegamos a string seguinte.
    '      Depois procuramos "fatorDiario" logo em seguida.
    '----------------------------------------------------------------------------
    
    Dim pos As Long
    pos = 1
    
    Dim rowOut As Long
    rowOut = 2
    
    Dim countFound As Long
    countFound = 0
    
    Dim posDataCotacao As Long
    Dim posDataValueStart As Long, posDataValueEnd As Long
    Dim txtDataCotacao As String
    
    Dim posFator As Long
    Dim txtFator As String
    
    ' Loop até não encontrar mais "dataCotacao":
    Do
        ' (a) Achar a próxima ocorrência de "dataCotacao":
        posDataCotacao = InStr(pos, jsonResponse, """dataCotacao"":", vbTextCompare)
        If posDataCotacao = 0 Then Exit Do  ' acabou
        
        '   Exemplo no JSON:
        '   "dataCotacao": "30/01/2025"
        '   Queremos pegar o texto entre as aspas após o :
        
        ' 1) Procurar a 1ª aspa (") depois de : 
        posDataValueStart = InStr(posDataCotacao, jsonResponse, """", vbTextCompare)
        If posDataValueStart = 0 Then Exit Do
        
        ' 2) Procurar a 2ª aspa (") que fecha a string
        posDataValueEnd = InStr(posDataValueStart + 1, jsonResponse, """", vbTextCompare)
        If posDataValueEnd = 0 Then Exit Do
        
        ' 3) Extrair a substring
        txtDataCotacao = Mid(jsonResponse, posDataValueStart + 1, posDataValueEnd - (posDataValueStart + 1))
        
        ' (b) Agora procurar por "fatorDiario" após posDataValueEnd
        posFator = InStr(posDataValueEnd, jsonResponse, """fatorDiario"":", vbTextCompare)
        If posFator = 0 Then Exit Do
        
        '   Exemplo no JSON:
        '   "fatorDiario": 1.00049037
        '   Precisamos pegar esse número até a próxima vírgula ou fecha-chaves
        txtFator = ExtrairValorNumerico(jsonResponse, posFator + Len("""fatorDiario"":"))
        
        ' (c) Preencher na planilha
        ws.Cells(rowOut, 1).Value = txtDataCotacao   ' DataCotacao (string)
        ws.Cells(rowOut, 1).NumberFormat = "dd/mm/yyyy"
        
        ws.Cells(rowOut, 2).Value = Val(txtFator)    ' fatorDiario (double)
        ws.Cells(rowOut, 2).NumberFormat = "#,##0.00000000"
        
        ws.Cells(rowOut, 3).Value = (rowOut - 1)     ' ID sequencial (1,2,3,...)
        
        rowOut = rowOut + 1
        countFound = countFound + 1
        
        ' Atualiza pos para próxima busca
        pos = posDataValueEnd + 1
    Loop
    
    MsgBox "Processo concluído! " & countFound & " linhas inseridas.", vbInformation
    
End Sub

'------------------------------------------------------------------------------
' Função auxiliar que, dado um trecho do JSON e uma posição,
' extrai um número (pode conter ponto ou vírgula decimal) até achar
' algum caractere que não faça parte do número (vírgula, fecha-chaves etc.).
'
' Exemplo no JSON:
'    "fatorDiario": 1.00049037,
'
' Chamamos: ExtrairValorNumerico(json, posDepoisDoDoisPontos)
'
' Ele ignora espaços e extrai "1.00049037".
'
' Obs: Se o JSON tiver formatação diferente, pode ser necessário ajustar.
'------------------------------------------------------------------------------
Private Function ExtrairValorNumerico(ByVal source As String, ByVal startPos As Long) As String
    
    Dim temp As String
    temp = Mid(source, startPos, 50)  ' pega até 50 chars a partir de startPos, ajuste se precisar
    
    temp = LTrim(temp)               ' remove espaços à esquerda
    
    Dim i As Long
    Dim ch As String
    Dim ret As String
    
    For i = 1 To Len(temp)
        ch = Mid(temp, i, 1)
        
        ' Aceitamos dígitos 0-9, ponto (.) ou vírgula (,) e o sinal de menos, se houver
        If (ch Like "[0-9]") Or (ch = ".") Or (ch = ",") Or (ch = "-") Then
            ret = ret & ch
        Else
            ' Paramos ao encontrar um caractere não aceito
            Exit For
        End If
    Next i
    
    ' Substituir vírgula por ponto (caso apareça no JSON 1,234567)
    ret = Replace(ret, ",", ".")
    
    ExtrairValorNumerico = ret
    
End Function
