Abaixo está um **exemplo completo** de como você pode:

1. **Limpar** a aba "selic2", da linha 2 até o fim (colunas que precisar).  
2. **Calcular** as datas de início (5 anos atrás) e final (hoje).  
3. **Enviar** um **POST** para o endpoint do BACEN.  
4. **Parsear** (analisar) o **JSON** de resposta.  
5. **Preencher** as colunas:
   - **A**: dataCotacao  
   - **B**: fatorDiario  
   - **C**: um ID sequencial começando em 1  

> **Observação sobre parsing de JSON**:  
> O VBA não traz nativamente um parser de JSON. Você pode usar:
> - [VBA-JSON (JsonConverter)](https://github.com/VBA-tools/VBA-JSON) do GitHub (Tim Hall), ou  
> - Alguma biblioteca interna de script (por exemplo, “Microsoft Script Control”) ou .NET, mas costuma ser mais simples importar o módulo **JsonConverter.bas** do VBA-JSON.
>
> Abaixo, assumo que você **já tem** o `JsonConverter` referenciado e pode fazer `Set json = JsonConverter.ParseJson(...)`.  

---

## Exemplo de Código

```vb
Public Sub ImportarSelic()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("selic2")

    '---------------------------------------------------------
    ' 1) APAGAR TUDO DA LINHA 2 ATÉ O FIM (COLUNAS A:C, por ex.)
    '    Ajuste o intervalo conforme a sua necessidade.
    '---------------------------------------------------------
    ws.Range("A2:C100000").ClearContents

    '---------------------------------------------------------
    ' 2) DEFINIR DATA DE HOJE E DATA INICIAL (5 ANOS ATRÁS)
    '---------------------------------------------------------
    Dim hoje As Date
    hoje = Date                     ' data atual do sistema

    Dim dtInicial As Date
    dtInicial = DateAdd("yyyy", -5, hoje)  ' subtrai 5 anos

    '---------------------------------------------------------
    ' 3) MONTAR O CORPO (BODY) DO POST EM FORMATO JSON
    '    O formato esperado é:
    '       { "dataInicial": "dd/mm/yyyy", "dataFinal": "dd/mm/yyyy" }
    '---------------------------------------------------------
    Dim postBody As String
    postBody = "{""dataInicial"":""" & Format(dtInicial, "dd/MM/yyyy") & """," & _
               """dataFinal"":""" & Format(hoje, "dd/MM/yyyy") & """}"

    '---------------------------------------------------------
    ' 4) FAZER A REQUISIÇÃO POST (USANDO MSXML2.XMLHTTP)
    '---------------------------------------------------------
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")

    With http
        .Open "POST", _
              "https://www3.bcb.gov.br/novoselic/rest/taxaSelicApurada/pub/search?parametrosOrdenacao=%5B%5D&page=1&pageSize=999999999", _
              False
        .setRequestHeader "Content-Type", "application/json"
        .send (postBody)
    End With

    '---------------------------------------------------------
    ' 5) PEGAR O JSON DE RESPOSTA
    '---------------------------------------------------------
    Dim jsonResponse As String
    jsonResponse = http.responseText

    '---------------------------------------------------------
    ' 6) PARSEAR O JSON (USANDO O VBA-JSON)
    '    Precisamos da função: Set json = JsonConverter.ParseJson(jsonResponse)
    '---------------------------------------------------------
    Dim json As Object
    Set json = JsonConverter.ParseJson(jsonResponse)

    ' A estrutura retorna algo como:
    ' {
    '   "totalItems": 21,
    '   "registros": [
    '       {"dataCotacao":"30/01/2025","fatorDiario":1.00049037,...},
    '       {"dataCotacao":"29/01/2025","fatorDiario":1.00045513,...},
    '        ...
    '   ]
    ' }
    '
    ' Então "registros" é um array/collection de objetos.

    Dim registros As Collection
    Set registros = json("registros")  ' Pega a coleção "registros"

    '---------------------------------------------------------
    ' 7) PREENCHER AS COLUNAS NA ABA: 
    '       A => dataCotacao
    '       B => fatorDiario
    '       C => id sequencial (1, 2, 3...)
    '---------------------------------------------------------
    Dim i As Long
    Dim row As Long
    row = 2  ' Começa a inserir na linha 2

    Dim item As Object
    
    For i = 1 To registros.Count
        Set item = registros(i)  ' Pega o i-ésimo registro

        ' A: dataCotacao
        ws.Cells(row, 1).Value = item("dataCotacao")
        ws.Cells(row, 1).NumberFormat = "dd/mm/yyyy"

        ' B: fatorDiario
        ws.Cells(row, 2).Value = item("fatorDiario")
        ws.Cells(row, 2).NumberFormat = "#,##0.00000000"

        ' C: ID sequencial
        ws.Cells(row, 3).Value = i

        row = row + 1
    Next i

    MsgBox "Processo concluído! " & (row - 2) & " linhas inseridas.", vbInformation

End Sub
```

### Passo a Passo do Código

1. **Limpar a aba**  
   - `ws.Range("A2:C100000").ClearContents` apaga qualquer conteúdo das colunas A até C, da linha 2 até a 100.000. Ajuste conforme necessário.

2. **Definir datas**  
   - `hoje = Date`: pega a data do sistema.  
   - `dtInicial = DateAdd("yyyy", -5, hoje)`: subtrai 5 anos.

3. **Criar o JSON body**  
   - Montamos a string `postBody` usando `Format(...)` em `"dd/MM/yyyy"` para respeitar o formato exigido pelo BACEN.

4. **Fazer a requisição HTTP**  
   - Usamos o objeto `MSXML2.XMLHTTP`.  
   - Método `Open "POST", url, False` para uma chamada síncrona (aguarda resposta).  
   - `setRequestHeader "Content-Type", "application/json"` informa que o corpo é JSON.  
   - `.send postBody` envia os dados.

5. **Pegar o JSON de resposta**  
   - Armazenado em `jsonResponse = http.responseText`.

6. **Parsear com VBA-JSON**  
   - `Set json = JsonConverter.ParseJson(jsonResponse)` converte a string JSON em um objeto (Dictionary/Collection).

7. **Extrair os “registros”**  
   - `Set registros = json("registros")` retorna a coleção (array) com cada dia/taxa.

8. **Inserir na planilha**  
   - Para cada item em `registros`, insere:
     - **Coluna A** (`dataCotacao`).  
     - **Coluna B** (`fatorDiario`).  
     - **Coluna C** (ID sequencial `i`).  
   - Aplica formatação de data ( `.NumberFormat = "dd/mm/yyyy"` ) e de número (`#,##0.00000000`, por exemplo).

---

### Requisitos

- **Referência** a alguma biblioteca de parsing JSON:  
  - Normalmente, adicionamos o módulo [**JsonConverter.bas** do repositório “VBA-JSON”](https://github.com/VBA-tools/VBA-JSON) no projeto VBA e marcamos a referência “Microsoft Scripting Runtime”.  
  - Ou seja, no VBA Editor: **Ferramentas > Referências** e marque “Microsoft Scripting Runtime” (se necessário).  
  - Importe o arquivo “JsonConverter.bas” no seu projeto.  

- **Conexão à internet** habilitada.  
- Caso a chamada HTTP encontre firewall ou proxy, pode ser necessário configurar.

---

Com esse código, você consegue:

- Limpar a área desejada.  
- Obter os dados de **até 5 anos atrás** (por default) até a data de hoje.  
- Listar cada dia/taxa (dataCotacao e fatorDiario) e um **ID sequencial** crescente.
